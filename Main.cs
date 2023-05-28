using System;
using System.Collections.Generic;
using HostMgd.ApplicationServices;
using Multicad.Runtime;
using Multicad.DatabaseServices;
using DocumentFormat.OpenXml.Spreadsheet;
using NCadCustom.Code;

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

        static readonly string[] excelExtentions = { ".xlsx", ".xls", ".xlsb", ".xlsm" };

        /// <summary>
        /// Создание СПДС объектов на чертеже по таблице из эксель
        /// Параметры и ID объектов находятся в эксель
        /// </summary>
        [CommandMethod("SpdsObjByExcelParams", CommandFlags.NoCheck | CommandFlags.NoPrefix)]
        public static void MainCreateObjBySpreadSheet()
        {
            InputJig jig = new InputJig();
            HostMgd.EditorInput.Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;


            string paramsFilePath = jig.GetText("Укажите полный путь до файла параметров(Excel)", false);
            paramsFilePath = paramsFilePath.Trim('"').ToLower();

            if (!File.Exists(paramsFilePath))
            {
                ed.WriteMessage("Выбран не существующий путь! Программа завершена.");
                return;
            }

            if (!excelExtentions.Contains(Path.GetExtension(paramsFilePath)))
            {
                ed.WriteMessage("Выбран не Excel файл! Программа завершена.");
                return;
            }

            try
            {
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

                ed.WriteMessage("Обработано!");
            }
            catch (Exception e)
            {
                ed.WriteMessage($"Ошибка : {e}");
            }
        }
    }
}
