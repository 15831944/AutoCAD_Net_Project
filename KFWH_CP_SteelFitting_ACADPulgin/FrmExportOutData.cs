using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using acadWnd = Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.DatabaseServices.Filters;

namespace KFWH_CP_SteelFitting_ACADPulgin
{
    public partial class FrmExportOutData : Form
    {
        public List<SteelFittingProfile> ListProfiles { get; set; }
        public FrmExportOutData()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.listBox1.Items.Clear();
            acadWnd.OpenFileDialog dlg = new Autodesk.AutoCAD.Windows.OpenFileDialog("选择图纸", null, "dwg", null, Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            if (dlg.ShowDialog() == DialogResult.OK) this.listBox1.Items.AddRange(dlg.GetFilenames().Where(c => MyUtility.FileInUse(c) == false).ToArray());
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (this.listBox1.Items.Count != 0)
            {
                this.ListProfiles = new List<SteelFittingProfile>();
                for (int i = 0; i < this.listBox1.Items.Count; i++)
                {
                    string str = this.listBox1.Items[i].ToString();
                    this.ListProfiles.AddRange(GetEachDrawingProfileData(str));
                }
                if (this.ListProfiles.Count > 0)
                {
                    Dictionary<string, double> dicList = new Dictionary<string, double>();
                    foreach (var item in this.ListProfiles)
                    {
                        if (dicList.ContainsKey(item.DwgNumber + "," + item.Size + "," + item.Grade))
                        {
                            dicList[item.DwgNumber + "," + item.Size + "," + item.Grade] += item.Length;
                        }
                        else dicList[item.DwgNumber + "," + item.Size + "," + item.Grade] = item.Length;
                    }
                    var dicWithSort = dicList.OrderBy(c => c.Key.Split(',')[0]).ThenBy(c => c.Key.Split(',')[1]).ThenBy(c => c.Key.Split(',')[2]).ToDictionary(c => c.Key, c => c.Value);
                    if (dicWithSort.Count > 0)
                    {
                        Excel.Application xlapp = new Excel.Application();
                        try
                        {
                            Excel.Workbook wb = xlapp.Workbooks.Add();
                            Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                            ws.Range["A1:D1"].Value = new string[] { "Drawing Number", "Size", "Grade", "Length" };
                            int k = 2;
                            foreach (KeyValuePair<string, double> item in dicWithSort)
                            {
                                ws.Range["A" + k + ":C" + k].Value = item.Key.Split(',');
                                ws.Range["D" + k].Value = item.Value;
                                k++;
                            }
                            ws.Range["A1:D1"].EntireColumn.AutoFit();
                            xlapp.Visible = true;
                            xlapp.WindowState = Excel.XlWindowState.xlMaximized;
                        }
                        catch (System.Exception ex)
                        {
                            xlapp.Quit();
                            MessageBox.Show(ex.Message);
                            throw;
                        }
                    }
                    else MessageBox.Show("Can't Find any Profile Information In Select Drawing !");
                }
                else MessageBox.Show("Can't Find any Profile Information In Select Drawing !");
            }
            else MessageBox.Show("请至少选择一个文件然后再执行导出操作！", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }
        private List<SteelFittingProfile> GetEachDrawingProfileData(string fileName)
        {
            List<SteelFittingProfile> profiles = new List<SteelFittingProfile>();
            using (Database db = new Database(false, true))
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                db.ReadDwgFile(fileName, System.IO.FileShare.Read, true, null);
                ////图层过滤
                //LayerFilter lyFilter = new LayerFilter();
                //lyFilter.Add("SteelFiting_ProfileLengthText");
                ////范围过滤
                //SpatialFilter spaFilter = new SpatialFilter();
                //Point2dCollection pts = new Point2dCollection();
                //pts.Add(db.Limmax);
                //pts.Add(db.Limmin);
                //spaFilter.Definition = new SpatialFilterDefinition(pts, Vector3d.ZAxis, 0, 0, 0, true);
                ////局部加载
                //db.ApplyPartialOpenFilters(spaFilter, lyFilter);
                ////关闭文件输入
                //if (db.IsPartiallyOpened)
                //{
                db.CloseInput(true);
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId item in btr)
                {
                    Entity ent = trans.GetObject(item, OpenMode.ForRead) as Entity;
                    if (ent is DBText && ent.Layer == "SteelFiting_ProfileLengthText")
                    {
                        if ((ent as DBText).TextString.Contains("=") && (ent as DBText).TextString.Contains(":")) profiles.Add(new SteelFittingProfile((ent as DBText).TextString));
                    }
                }
                //}
            }
            return profiles;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
    public class SteelFittingProfile
    {
        public double Length { get; set; }
        public string Size { get; set; }
        public string Grade { get; set; }
        public string DwgNumber { get; set; }
        public SteelFittingProfile(string Textstring)
        {
            //H33001:FB 6x1/4"(AH36)=537.46mm
            string[] arrs = Textstring.Split(new char[] { ':', '(', ')', '=' }, StringSplitOptions.RemoveEmptyEntries);
            switch (arrs.Length)
            {
                case 3:
                    this.DwgNumber = arrs[0];
                    this.Size = arrs[1];
                    this.Length = double.Parse(arrs[2].Replace("mm", "").Trim());
                    this.Grade = "N.A.";
                    break;
                case 4:
                    this.DwgNumber = arrs[0];
                    this.Size = arrs[1];
                    this.Grade = arrs[2];
                    this.Length = double.Parse(arrs[3].Replace("mm", "").Trim());
                    break;
                case 5:
                    this.DwgNumber = arrs[0];
                    this.Size = arrs[1];
                    this.Grade = arrs[2];
                    this.Length = double.Parse(arrs[4].Replace("mm", "").Trim());
                    break;
            }
            if (this.Size.Contains("%%C")) this.Size = this.Size.Replace("%%C", "ø ");
            if (this.Size.Contains("%%D")) this.Size = this.Size.Replace("%%D", "°");
        }
    }
}
