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

        }
    }
}
