using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using static System.IO.File;

namespace ExcelAlerter
{
    public enum Mode
    {
        Column = 1,
        Row = 2
    }
    public interface IAlerter
    {
        List<Engineer> Data { get; }
        void Notify();
        void LoadData();
    }
    public class Alerter : IAlerter
    {
        private readonly ILogger<Alerter> _log;
        private readonly AppSettings _settings;
        private readonly IHostingEnvironment _env;
        public List<Engineer> Data { get; private set; } = new List<Engineer>();

        public Alerter(ILogger<Alerter> log,
                       IOptionsMonitor<AppSettings> settings,
                       IHostingEnvironment env)
        {
            _log = log;
            _settings = settings.CurrentValue;
            _env = env;
        }

        public void LoadData()
        {
            string absolutePath = $"{_env.ContentRootPath}{Path.DirectorySeparatorChar}{_settings.FilePath}";

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(absolutePath, true))
            {
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(f => _settings.Worksheets.ToList().Contains(f.Name));
                if (sheets.Count() == 0)
                {
                    // The specified worksheet does not exist.
                    _log.LogWarning("Worksheets not found!");
                    return;
                }

                //TODO: need to loop here

                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData worksheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                WorkbookPart workbookPart = document.WorkbookPart;

                GetEngineerSchedule(worksheet, worksheetData, workbookPart, "Engineer");
                GetEngineerSchedule(worksheet, worksheetData, workbookPart, "Team Leader");
            }

        }

        private void GetEngineerSchedule(Worksheet worksheet, SheetData worksheetData, WorkbookPart workbookPart, string role)
        {
            var engineerCells = FindCellsWithText(workbookPart, worksheetData, $"{role}:");

            _log.LogInformation($"Found {engineerCells.Count()} Engineers{Environment.NewLine}");

            //Engineer Cell is top left of a work item
            //It is 3 columns wide and N columns deep
            //Cell + 1 row is Equipment Header
            //Cell + 1 column + 1 row is ID Header
            //Cell + 1 column + 2 rows is Date Header
            //Cell + N rows until empty/or red is Equipment Names
            //Cell + 1 column + N rows until empty/or red is IDs
            //Cell + 2 column + N rows until empty/or red is Dates
            foreach (Cell engineerCell in engineerCells)
            {
                try
                {
                    Engineer engineer = AddSingleEngineerSchedule(worksheet, workbookPart, role, engineerCell);

                    Data.Add(engineer);    
                }
                catch (Exception ex)
                {
                    _log.LogError($"Failed to Load Engineer Data. {ex.ToString()}");
                }
                
            }
        }

        private Engineer AddSingleEngineerSchedule(Worksheet worksheet, WorkbookPart workbookPart, string role, Cell engineerCell)
        {
            var engineerName = GetCellValue(workbookPart, engineerCell).Replace($"{role}:", "").Trim();

            var engineer = Engineer.Create(engineerName, role);

            _log.LogInformation("");
            _log.LogInformation($"{Environment.NewLine}Processing {engineerName} ({role})");

            var eqipmentHeaderCellref = Increment(engineerCell.CellReference.Value, Mode.Row);

            //get first piece of equipment
            var eqiupmentCellref = eqipmentHeaderCellref;
            while (true)
            {

                //TODO - Handle invalid or empty cells

                eqiupmentCellref = Increment(eqiupmentCellref, Mode.Row);
                Cell equipmentCell = GetCellUsingReference(worksheet, eqiupmentCellref);
                var equipmentName = GetCellValue(workbookPart, equipmentCell);

                if(string.IsNullOrWhiteSpace(equipmentName))
                    break;

                var idCellref = Increment(eqiupmentCellref, Mode.Column);
                Cell idCell = GetCellUsingReference(worksheet, idCellref);
                var id = GetCellValue(workbookPart, idCell);

                var dateCellref = Increment(idCellref, Mode.Column);
                Cell dateCell = GetCellUsingReference(worksheet, dateCellref);
                
                var dateText = GetCellValue(workbookPart, dateCell);
                var date = DateTime.MinValue;
                
                if(double.TryParse(dateText, out double dateDouble))
                {
                    date = DateTime.FromOADate(dateDouble);

                    _log.LogInformation($"Added {equipmentName} - {id} - {date.ToShortDateString()}");
                }
                else
                {
                    _log.LogInformation($"Added {equipmentName} - {id} - No Date Specificed");
                }

                engineer.AddEquipment(Equipment.Create(id, equipmentName, date));
            }

            return engineer;
        }

        private static Cell GetCellUsingReference(Worksheet worksheet, string cellRef)
        {
            return worksheet.Descendants<Cell>().Where(c => c.CellReference == cellRef).FirstOrDefault();
        }

        public List<Cell> FindCellsWithText(WorkbookPart wb, SheetData sheetData, string textToFind)
        {
            var arr = new List<Cell>();
            string values;
           
            foreach(Row r in sheetData.Elements<Row>())
            {
                foreach (Cell c in r.Elements<Cell>())
                {
                    if(GetCellValue(wb, c).Contains(textToFind))
                        arr.Add(c);
                }
            }
            return arr;
        }
        public List<string> SearchCellsText(WorkbookPart wb, SheetData sheetData, string textToFind)
        {
            var arr = new List<string>();
            string values;
           
            foreach(Row r in sheetData.Elements<Row>())
            {
                values = "";
                foreach (Cell c in r.Elements<Cell>())
                {
                    values += GetCellValue(wb, c) + ",";
                }
                if (values.Length > 0)
                    values = values.Substring(0, values.Length - 1);

                if(values.Contains(textToFind))
                    arr.Add(values);
            }
            return arr;
        }

        public string GetCellValue(WorkbookPart wb, Cell c)
        {
            if(c == null)return "";

            string value = c.InnerText;
            if (c.DataType != null)
            {
                switch (c.DataType.Value)
                {
                    case CellValues.SharedString:
                        var stringTable = wb.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (stringTable != null)
                        {
                            value =
                                stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                    default:
                        value = "";
                        break;
                }
            }
            return value;
        }

        public static string Increment(string text, Mode mode)
        {
            Regex re = new Regex(@"([a-zA-Z]+)(\d+)");
            Match result = re.Match(text);

            char column = result.Groups[1].Value.ToCharArray().First(); //A,B,C etc - We assume we never get past Z here
            int row = Convert.ToInt16(result.Groups[2].Value); //1, 2, 3 etc

            if(mode == Mode.Column)
                column = (Char)(Convert.ToUInt16(column) + 1);

            if(mode == Mode.Row)
                row++;

            return $"{column}{row}";
        }


        public void Notify()
        {
            _log.LogInformation("Notifying");
        }
    }
}