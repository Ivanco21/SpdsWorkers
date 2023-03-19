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

namespace SpdsObjBySheetInfo
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
        public void MainCreateBySpreadSheet()
        {
            long idEl = 5252373799596257837;
            //McParametricObject mcParametric = new McParametricObject();
            //bool res = mcParametric.Initialize(idEl); //519ACE905DC70EB0

            //Document dbDoc = new Document(true, true);

            //Folder fldr = new Folder(;

            Connection conn = new Connection();
            //ElementFilter filter = new ElementFilter();

            //filter.IncludeSubfolders = true;
            //filter

            Element el = conn.GetElement(idEl);

            int test = 5;

        }
    }
}
