using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using System.Reflection;

namespace KFWH_CMD
{
    public class AutoCADNetLoader
    {
        [CommandMethod("MyNetLoader")]
        public void Dll加载到内存并卸载()
        {
            if (Environment.UserName != "sheng.nan"&& Environment.UserName != "peng.shen" && Environment.UserName != "yaofan.liu") { Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("你无权使用该工具！");return; }
            Autodesk.AutoCAD.Windows.OpenFileDialog ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog("Dll加载到内存并卸载", "Dll加载到内存并卸载", "dll", "Dll加载到内存并卸载", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (var item in ofd.GetFilenames())
                {
                    var assBytes=System.IO.File.ReadAllBytes(item);
                    var ass=Assembly.Load(assBytes);
                }
            }
            else Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("请选择一个或者多个dll文件！");
        }
    }
}
