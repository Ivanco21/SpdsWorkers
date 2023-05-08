using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpdsObjBySheetInfo
{
    internal class ShWorker
    {
        private readonly string wbFullPath;
        internal List<Row> dataRows;

        public ShWorker(string shPath) 
        {
            this.wbFullPath = shPath;
            this.dataRows = ReadDocRows();
        }

        private List<Row> ReadDocRows()
        {
            List<Row> resRows = new List<Row>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(wbFullPath, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                WorksheetPart shPart = wbPart.WorksheetParts.First();

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
