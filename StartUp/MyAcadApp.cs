using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Microsoft.CSharp;
using System.IO;

//
using Autodesk.AutoCAD.Windows;
using Autodesk.Windows;
using System.Reflection;
using System.Data.SqlClient;
using System.Security.AccessControl;
using System.Diagnostics;
using Autodesk.AutoCAD.Customization;
using System.Collections.Specialized;

//[assembly: ExtensionApplication(typeof(KFWH_CMD.MyAcadApp))]
namespace KFWH_CMD
{
    public class MyAcadApp : IExtensionApplication
    {
        public string mycuiFileName { get; set; }
        public string currentDllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        public void Initialize()//初始化程序。
        {
            AcadNetDllAutoLoader(Registry.CurrentUser);                //写入注册表
            var netDllList = AutoCADNetDllInfor();
            List<string> localDlllist = new List<string>();
            // 检测本地的dll 并从服务器下载文件
            foreach (string dllFileName in netDllList)
            {
                FileInfo fileCurrent = new FileInfo(currentDllPath + "\\" + Path.GetFileNameWithoutExtension(dllFileName));
                FileInfo fileNet = new FileInfo(dllFileName);
                if (fileNet.Exists) fileNet.CopyTo(currentDllPath + "\\" + Path.GetFileName(dllFileName), true);
                localDlllist.Add(currentDllPath + "\\" + Path.GetFileName(dllFileName));
            }
            //复制图标
            #region// copy cmd icon to local folder
            if (!Directory.Exists(currentDllPath + "\\KFWH_CMD_ICON")) Directory.CreateDirectory(currentDllPath + "\\KFWH_CMD_ICON");
            //
            List<string> listIconFiles = new List<string>();
            string sqlcommand = $"select * from tb_PluginUsedFiles where Class_Name='AutoCADPluginICON' and FileName='KFWH_CMD_ICON'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    listIconFiles.Add(sdr["Path"].ToString() + "\\" + sdr["FileName"].ToString());
                }
            }
            if (Directory.Exists(listIconFiles[0]))
            {
                string[] fis = Directory.GetFiles(listIconFiles[0]);
                foreach (var item in fis)
                {
                    FileInfo fi = new FileInfo(item);
                    fi.CopyTo(currentDllPath + "\\KFWH_CMD_ICON\\" + Path.GetFileName(item), true);
                }
            }
            #endregion
            //加载dll到cad,提取cad命令
            List<myAutoCADNetDll> listDllcmd = GetACADNetDllCmd(localDlllist);
            foreach (var item in localDlllist) ExtensionLoader.Load(item);
            AddMenusToAutoCAD(listDllcmd);//加载cuix菜单
            if (this.mycuiFileName != string.Empty) Application.LoadPartialMenu(this.mycuiFileName);
        }

        private void AcadApp_EndOpen(string FileName)
        {
            Application.SetSystemVariable("MENUBAR", 1);
            Application.SetSystemVariable("NAVBARDISPLAY", 1);
        }
        public void Terminate()
        {
            //AutoCADNetPluginMainDll
            try
            {
                //删除旧的dll
                FileInfo fi = new FileInfo(@"D:\CsharpPulgin\" + "KFWH_CP_PlugIn.dll"); if (fi.Exists) fi.Delete();
                FileInfo fi1 = new FileInfo(@"D:\CsharpPulgin\" + "MYCADCMD.dll"); if (fi1.Exists) fi1.Delete();
                FileInfo fi2 = new FileInfo(@"D:\CsharpPulgin\" + "StartUp.dll"); if (fi2.Exists) fi2.Delete();
                FileInfo fi3 = new FileInfo(@"D:\CsharpPulgin\" + "KFWH_CP_SteelFitting_ACADPlugin.dll"); if (fi3.Exists) fi3.Delete();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            finally
            {
                if (this.mycuiFileName != string.Empty)
                {
                    Application.UnloadPartialMenu(this.mycuiFileName);
                    File.Delete(this.mycuiFileName.Replace(".cuix", ".mnr"));
                    File.Delete(this.mycuiFileName.Replace(".cuix", "_light.mnr"));
                    File.Delete(this.mycuiFileName);
                }
            }
        }
        //20190318

        /// <summary>
        /// 写入程序到注册表，实现加载一次后，自动加载到AutoCAD
        /// </summary>
        /// <param name="reg_Key">加入到那个注册表，HKLM?,HKCU?</param>
        public void AcadNetDllAutoLoader(RegistryKey reg_Key)
        {
            string regPath = ""; bool flag_currentApp = true; bool flag_oldApp = false;
            if (reg_Key.ToString() == Registry.LocalMachine.ToString())
            {
                regPath = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.MachineRegistryProductRootKey;
            }
            else regPath = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
            string assemblyFileName = Assembly.GetExecutingAssembly().CodeBase;
            RegistryKey acad_key = reg_Key.OpenSubKey(Path.Combine(regPath, "Applications"), false);
            foreach (var item in acad_key.GetSubKeyNames()) if (item == Path.GetFileNameWithoutExtension(assemblyFileName)) flag_currentApp = false;
            if (flag_currentApp)
            {
                acad_key = reg_Key.OpenSubKey(Path.Combine(regPath, "Applications"), true);
                RegistryKey myAppkey = acad_key.CreateSubKey(Path.GetFileNameWithoutExtension(assemblyFileName), Microsoft.Win32.RegistryKeyPermissionCheck.Default);
                myAppkey.SetValue("DESCRIPTION", "加载自定义Dll");
                myAppkey.SetValue("LOADCTRLS", 2, Microsoft.Win32.RegistryValueKind.DWord);
                myAppkey.SetValue("LOADER", assemblyFileName, Microsoft.Win32.RegistryValueKind.String);
                myAppkey.SetValue("MANAGED", 1, Microsoft.Win32.RegistryValueKind.DWord);
                Application.ShowAlertDialog($"{Path.GetFileNameWithoutExtension(assemblyFileName)} 程序加载完成，重启CAD生效！");
            }
            else Application.ShowAlertDialog($"欢迎使用 {Path.GetFileNameWithoutExtension(assemblyFileName)} 程序！");
            #region//删除旧的程序
            foreach (var item in acad_key.GetSubKeyNames()) if (item == "KFWH_CP_PlugIn") flag_oldApp = true;
            if (flag_currentApp)
            {
                acad_key = reg_Key.OpenSubKey(Path.Combine(regPath, "Applications"), true);
                acad_key.DeleteSubKey("KFWH_CP_PlugIn");
                Application.ShowAlertDialog($"请重启CAD以完成旧程序的删除");
            }
            #endregion
        }
        /// <summary>
        /// 通过反射机制导出AutoCAD Net Dll的注册的命令
        /// </summary>
        /// <param name="ass">包含AutoCAD Net Dll的注册的命令的Dll</param>
        /// <returns></returns>
        public List<myAutoCADNetDll> GetACADNetDllCmd(List<string> dllFileNames)
        {
            List<myAutoCADNetDll> listDlls = new List<myAutoCADNetDll>();
            #region//收集自定义命令
            try
            {
                foreach (var item in dllFileNames)
                {
                    Assembly ass = Assembly.LoadFrom(item);
                    List<myAutoCADNetCmd> ListCmds = new List<myAutoCADNetCmd>();
                    foreach (var t in ass.GetTypes())
                    {
                        if (t.IsClass && t.IsPublic)
                        {
                            foreach (MethodInfo method in t.GetMethods())
                            {
                                if (method.IsPublic && method.GetCustomAttributes(true).Length > 0)
                                {
                                    CommandMethodAttribute cadAtt = null; myAutoCADcmdAttribute myAtt = null;
                                    foreach (var att in method.GetCustomAttributes(true))
                                    {
                                        if (att.GetType().Name == typeof(CommandMethodAttribute).Name) cadAtt = att as CommandMethodAttribute;
                                        if (att.GetType().Name == typeof(myAutoCADcmdAttribute).Name) myAtt = att as myAutoCADcmdAttribute;
                                    }
                                    if (myAtt != null && cadAtt != null)
                                    {
                                        ListCmds.Add(new myAutoCADNetCmd(Path.GetFileNameWithoutExtension(
                                            ass.ManifestModule.Name.Substring(0, ass.ManifestModule.Name.Length - 4)),
                                            t.Name, method.Name, cadAtt.GlobalName, myAtt.HelpString, myAtt.IcoFileName));
                                    }
                                    else
                                    {
                                        if (cadAtt != null && myAtt == null)
                                        {
                                            ListCmds.Add(new myAutoCADNetCmd(Path.GetFileNameWithoutExtension(
                                                ass.ManifestModule.Name.Substring(0, ass.ManifestModule.Name.Length - 4)),
                                                t.Name, method.Name, cadAtt.GlobalName, "", ""));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (ListCmds.Count > 0)
                    {
                        listDlls.Add(new myAutoCADNetDll(ass, ListCmds));
                    }
                }
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog(ex.Message);
            }
            #endregion
            return listDlls;
        }
        /// <summary>
        /// 将自定义命令定义成菜单
        /// </summary>
        /// <param name="dic"></param>
        public void AddMenusToAutoCAD(List<myAutoCADNetDll> dicCmds)
        {
            string strCuiFileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\KFWH_CMD.cuix";
            string strGpName = "KFWH_CMDGroup";//defined a group name
            CustomizationSection myCsec = new CustomizationSection() { MenuGroupName = strGpName };
            MacroGroup mg = new MacroGroup("KFWH_CMD", myCsec.MenuGroup);
            StringCollection scMyMenuAlias = new StringCollection();
            scMyMenuAlias.Add("KFWH_CMD_POP");
            PopMenu pmParent = new PopMenu("KFWH_CMD", scMyMenuAlias, "KFWH_CMD", myCsec.MenuGroup);
            foreach (var dll in dicCmds)//each assembly commmands
            {
                PopMenu pmCurDll = new PopMenu(dll.DllName, new StringCollection(), "", myCsec.MenuGroup);
                PopMenuRef pmrdll = new PopMenuRef(pmCurDll, pmParent, -1);
                foreach (var cls in dll.ListsNetClass)
                {
                    if (cls.hasCmds)
                    {
                        PopMenu pmCurcls = new PopMenu(cls.className, new StringCollection(), "", myCsec.MenuGroup);
                        PopMenuRef pmrcls = new PopMenuRef(pmCurcls,pmCurDll, -1);
                        foreach (var cmd in cls.ListsNetcmds)
                        {
                            MenuMacro mm = new MenuMacro(mg, cmd.MethodName, cmd.cmdName, cmd.DllName+"_"+cmd.className+"_"+cmd.cmdName, MacroType.Any);
                            //指定命令宏的说明信息，在状态栏中显示
                            if (cmd.HelpString!=string.Empty) mm.macro.HelpString = cmd.HelpString;
                            if (cmd.IcoName != string.Empty) mm.macro.LargeImage = mm.macro.SmallImage = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\KFWH_CMD_ICON\\"+cmd.IcoName;
                            //指定命令宏的大小图像的路径
                            PopMenuItem newPmi = new PopMenuItem(mm, cmd.MethodName, pmCurcls, -1);
                            //newPmi.MacroID = cmd.DllName + "_" + cmd.className + "_" + cmd.cmdName;
                        }
                    }
                }
            }
            //myCsec.MenuGroup.AddToolbar("kfwh_cmd");
            if (myCsec.SaveAs(strCuiFileName)) this.mycuiFileName = strCuiFileName;
        }
        public List<string> AutoCADNetDllInfor()
        {
            List<string> listDlls = new List<string>();
            string sqlcommand = $"select * from tb_PluginUsedFiles where Class_Name='Dll_AutoCAD'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    listDlls.Add(sdr["Path"].ToString() + "\\" + sdr["FileName"].ToString());
                }
            }
            return listDlls;
        }
    }


    public class myAutoCADNetCmd
    {
        public string DllName { get; set; }
        public string className { get; set; }
        public string MethodName { get; set; }
        public string cmdName { get; set; }
        public string HelpString { get; set; }
        public string IcoName { get; set; }
        public myAutoCADNetCmd(string _dllName, string _clsName, string _meName, string _cmdName, string _helpStr, string _icoName)
        {
            this.DllName = _dllName;
            this.className = _clsName;
            this.MethodName = _meName;
            this.cmdName = _cmdName;
            this.HelpString = _helpStr;
            this.IcoName = _icoName;
        }
    }

    /// <summary>
    /// 按程序集
    /// </summary>
    public class myAutoCADNetDll
    {
        public string DllName { get; set; }
        public List<myAutoCADNetClass> ListsNetClass { get; set; }
        public myAutoCADNetDll(Assembly ass, List<myAutoCADNetCmd> ListCmds)
        {
            this.ListsNetClass = new List<myAutoCADNetClass>();
            this.DllName = ass.ManifestModule.Name.Substring(0, ass.ManifestModule.Name.Length - 4);
            var cmdsByDll = ListCmds.Where(c => c.DllName == this.DllName).ToList();
            List<string> clsNames = new List<string>();
            foreach (myAutoCADNetCmd item in cmdsByDll) if (!clsNames.Contains(item.className)) clsNames.Add(item.className);

            foreach (string clsName in clsNames)
            {
                var netCls = new myAutoCADNetClass(this.DllName, clsName, cmdsByDll);
                if (netCls.hasCmds) netCls.GetListCmds(cmdsByDll);
                if (netCls.ListsNetcmds != null) ListsNetClass.Add(netCls);
            }
        }
    }

    /// <summary>
    /// 程序集下按类分类命令
    /// </summary>
    public class myAutoCADNetClass
    {
        public string className { get; set; }
        public string dllName { get; set; }
        public List<myAutoCADNetCmd> ListsNetcmds { get; set; }
        public bool hasCmds { get; set; }
        public myAutoCADNetClass(string _dllName, string _clsName, List<myAutoCADNetCmd> ListCmds)
        {
            this.className = _clsName;
            this.dllName = _dllName;
            if (ListCmds.Count(c => c.DllName == this.dllName && c.className == this.className) > 0) hasCmds = true; else hasCmds = false;
        }
        public void GetListCmds(List<myAutoCADNetCmd> ListCmds)
        {
            if (this.hasCmds)
            {
                this.ListsNetcmds = ListCmds.Where(c => c.DllName == this.dllName && c.className == this.className).ToList();
            }
            else this.ListsNetcmds = null;
        }
    }

}
