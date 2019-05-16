using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Windows;
using KFWH_CMD;
using System.Data.SqlClient;

namespace KFWH_CP_SD
{
    public class SetUp
    {
        [myAutoCADcmd("Update shop drawing title blcok information as per panel what you select", "SD_PanelInfo.png")]
        [CommandMethod("SD_PanelInfo")]
        public void KFWHSD_设置Panel信息()
        {
            PanelInformationSetup frm = new PanelInformationSetup();
            Application.ShowModelessDialog(frm);
        }
    }
}
