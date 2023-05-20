using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using Multicad.Geometry;
using Multicad;
using Multicad.Objects;

namespace SpdsObjBySheetInfo
{
    internal class ObjectForInsert
    {
        internal ObjectForInsert(SharedStringTable sst, List<string> headers, Row objectData)
        {
            CreateObject(sst, headers, objectData);
        }

        internal string dataBaseID { get; set; } //hexadecimal string 16
        internal int detailCount { get; set; }
        internal double x_coord { get; set; }
        internal double y_coord { get; set; }
        internal Dictionary<string, double> objParams { get; set; }


        private void CreateObject(SharedStringTable sst, List<string> headers, Row objectData)
        {
            // парсинг - по номерам столбцов
            IEnumerable<Cell> cells = objectData.Elements<Cell>();
            this.objParams = new Dictionary<string, double>();

            Cell cl = cells.ElementAt(0);
            if ((cl.DataType != null) && (cl.DataType == CellValues.SharedString))
            {
                int ssid = int.Parse(cl.CellValue.Text);
                string cellValue = sst.ChildElements[ssid].InnerText;
                this.dataBaseID = cellValue;
            }

            int intParsedValue = 0;
            if (int.TryParse(cells.ElementAt(1).CellValue.Text, out intParsedValue))
            {
                this.detailCount = intParsedValue;
            }

            // X,Y coords parse
            double doubleParsedValue = 0;
            if (double.TryParse(cells.ElementAt(2).CellValue.Text, out doubleParsedValue))
            {
                this.x_coord = doubleParsedValue;
            }
            if (double.TryParse(cells.ElementAt(3).CellValue.Text, out doubleParsedValue))
            {
                this.y_coord = doubleParsedValue;
            }

            // начинаются параметры для объектов БД 
            for (int iCol = 4; iCol < cells.Count(); iCol++)
            {
                doubleParsedValue = 0;
                if (double.TryParse(cells.ElementAt(iCol).CellValue.Text, out doubleParsedValue))
                {
                   objParams.Add(headers[iCol], doubleParsedValue);
                }
            }
        }

        internal void PlaceToModelSpace()
        {
            McParametricObject parObj = new McParametricObject(false);
            parObj.DbEntity.AddToCurrentDocument();

            long id = Convert.ToInt64(this.dataBaseID, 16);
            parObj.Initialize(id);
            parObj.SetImplementationAndProcess("Implementation 1");
            parObj.ViewInXY = McParametricObject.ViewType.CodirectX;

            List<ExValue> paramsToChange = new List<ExValue>();
            foreach (KeyValuePair<string, double> paramsPair in this.objParams)
            {
                paramsToChange.Add(new ExValue(paramsPair.Key, paramsPair.Value));
            }

            parObj.Change(paramsToChange, true);

            Matrix3d tfm = Matrix3d.MakeTranslation(this.x_coord, this.y_coord, 0);
            parObj.DbEntity.Transform(tfm);
        }

    }
}
