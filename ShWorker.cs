using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpdsObjBySheetInfo
{
    internal class ShWorker
    {
        private readonly string wbFullPath;
        private readonly string sheetName;

        public ShWorker(string shPath , string sheetName) 
        {
            this.wbFullPath = shPath;
            this.sheetName = sheetName;
            ReadDoc();
        }

        private void ReadDoc()
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(wbFullPath, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                Sheet sh = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

                if (sh == null)
                {
                    throw new ArgumentException(sheetName);
                }

                Row row = sh.Descendants<Row>().LastOrDefault();


            }
        }
    }
}
