using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NCadCustom.Code
{
    internal class ShWorker
    {
        private readonly string wbFullPath;
        internal List<Row> dataRows;
        internal SharedStringTable sst; // объект необходим чтобы корректно читать строки в Open Xml

        public ShWorker(string shPath)
        {
            wbFullPath = shPath;
            dataRows = ReadDocRows();
        }

        /// <summary>
        /// Получение списка с названиям колонок - headers
        /// </summary>
        /// <param name="sst"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        internal List<string> GetHeaders(SharedStringTable sst, Row row)
        {
            List<string> headers = new List<string>();

            foreach (var cl in row.Elements<Cell>())
            {
                if (cl.DataType != null && cl.DataType == CellValues.SharedString)
                {
                    int ssid = int.Parse(cl.CellValue.Text);
                    string cellValue = sst.ChildElements[ssid].InnerText;
                    headers.Add(cellValue);
                }
            }

            return headers;
        }

        /// <summary>
        /// чтение документа эксель
        /// </summary>
        /// <returns>массив строк эксель</returns>
        /// <exception cref="ArgumentException"></exception>
        private List<Row> ReadDocRows()
        {
            List<Row> resRows = new List<Row>();
            using (var fileStream = new FileStream(wbFullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileStream, false))
                {
                    WorkbookPart wbPart = doc.WorkbookPart;
                    WorksheetPart shPart = wbPart.WorksheetParts.First();
                    SharedStringTablePart sstpart = wbPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;
                    this.sst = sst;

                    if (shPart == null)
                    {
                        throw new ArgumentException();
                    }

                    SheetData sheetData = shPart.Worksheet.Elements<SheetData>().First();

                    Cell controlCell = null;
                    CellValue controlValue = null;
                    string controlText = string.Empty;
                    foreach (Row r in sheetData.Elements<Row>())
                    {
                        controlCell = r.Elements<Cell>().First();
                        if (controlCell == null)
                        {
                            break;
                        }

                        controlValue = controlCell.CellValue;
                        if (controlValue == null)
                        {
                            break;
                        }

                        controlText = controlValue.Text;

                        if (controlText == string.Empty)
                        {
                            break;
                        }

                        resRows.Add(r);
                    }
                }
            }

            return resRows;
        }
    }
}
