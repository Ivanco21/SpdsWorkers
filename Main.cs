using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime;
using Multicad.Geometry;
using Multicad.Runtime;
using Multicad.Symbols.Tables;
using Multicad.DataServices;
using Multicad.Objects;
using SpdsObjBySheetInfo;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NCadCustom
{
    public class Commands : IExtensionApplication
    {
        public void Initialize()
        {
        }
        public void Terminate()
        {
        }

        [CommandMethod("CreateSpdsBySh", CommandFlags.NoCheck | CommandFlags.NoPrefix)]
        public static void MainCreateBySpreadSheet()
        {
            string paramsFilePath = @"D:\\Code_repository\\C#\\nanoCAD\\SpdsObjBySheetInfo\\TestFiles\\Params.xlsm";
            ShWorker shWorker = new ShWorker(paramsFilePath);
            List<Row> dataRows = shWorker.dataRows;

            List<string> headers = shWorker.GetHeaders(shWorker.sst, dataRows.ElementAt(0));
            List<ObjectForInsert> objs = new List<ObjectForInsert>();

            // за исключением шапки - остальное строки с данными о деталях.
            Row rw = new Row();
            for (int iRow = 1; iRow < dataRows.Count; iRow++)
            {
                rw = dataRows[iRow];
                ObjectForInsert oneObject = new ObjectForInsert(shWorker.sst, headers, rw);
                objs.Add(oneObject);
            }

            foreach (ObjectForInsert obj in objs)
            {
                obj.PlaceToModelSpace();
            }



        }
    }
}
