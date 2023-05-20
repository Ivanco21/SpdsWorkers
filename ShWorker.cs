using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpdsObjBySheetInfo
{
    internal class ShWorker
    {
        private readonly string wbFullPath;
        internal List<Row> dataRows;
        internal SharedStringTable sst; // объект необходим чтобы корректно читать строки в Open Xml

        public ShWorker(string shPath)
        {
            this.wbFullPath = shPath;
            this.dataRows = ReadDocRows();
        }

        internal List<string> GetHeaders(SharedStringTable sst, Row row)
        {
            List<string> headers = new List<string>();

            foreach (var cl in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                if ((cl.DataType != null) && (cl.DataType == CellValues.SharedString))
                {
                    int ssid = int.Parse(cl.CellValue.Text);
                    string cellValue = sst.ChildElements[ssid].InnerText;
                    headers.Add(cellValue);
                }
            }

            return headers;
        }


        private List<Row> ReadDocRows()
        {
            List<Row> resRows = new List<Row>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(wbFullPath, false))
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
            return resRows;
        }
    }
}
