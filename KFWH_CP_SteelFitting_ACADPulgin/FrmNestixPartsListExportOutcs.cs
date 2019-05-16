using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// autocad api
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using acadWnd = Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using KFWH_CP_SteelFitting_ACADPulgin;

namespace KFWH_CP_SteelFitting_ACADPulgin
{
    public partial class FrmNestixPartsListExportOutcs : Form
    {
        public List<string[]> listNcParts { get; set; }
        public FrmNestixPartsListExportOutcs()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.listBox1.Items.Clear();
            acadWnd.OpenFileDialog dlg = new Autodesk.AutoCAD.Windows.OpenFileDialog("选择图纸", null, "dwg;dxf", null, Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowMultiple);
            if (dlg.ShowDialog() == DialogResult.OK) this.listBox1.Items.AddRange(dlg.GetFilenames().Where(c => MyUtility.FileInUse(c) == false).ToArray());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.listBox1.Items.Count != 0)
            {
                this.listNcParts = new List<string[]>();
                for (int i = 0; i < this.listBox1.Items.Count; i++)
                {
                    string str = this.listBox1.Items[i].ToString();
                    this.listNcParts.AddRange(GetEachDrawingNestixPartsData(str));
                }
                if (this.listNcParts.Count > 0)
                {
                    Excel.Application xlapp = new Excel.Application();
                    try
                    {
                        Excel.Workbook wb = xlapp.Workbooks.Add();
                        Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                        ws.Range["A1:G1"].Value = new string[] { "Name", "Thickness", "Grade", "Qty", "SIDE", "MirrorQty", "FileName" };
                        int k = 2;
                        foreach (var item in this.listNcParts)
                        {
                            ws.Range["A" + k + ":G" + k].Value = item;
                            k++;
                        }
                        ws.Range["A1:G1"].EntireColumn.AutoFit();
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
            }
            else MessageBox.Show("请至少选择一个文件再执行导出操作！", "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }

        private List<string[]> GetEachDrawingNestixPartsData(string fileName)
        {
            List<string[]> listPartInfor = new List<string[]>();
            List<DBText> listText = new List<DBText>();
            using (Database db = new Database(false, true))
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                db.ReadDwgFile(fileName, System.IO.FileShare.Read, true, null);
                db.CloseInput(true);
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                foreach (ObjectId item in btr)
                {
                    Entity ent = trans.GetObject(item, OpenMode.ForRead) as Entity;
                    if (ent is DBText && ent.Layer == "NC-LABEL") listText.Add(ent as DBText);
                }
                for (int i = listText.Count - 1; i >= 0; i--)
                {
                    string[] str = new string[7];
                    var item = listText[i];
                    if (item.TextString.StartsWith("N:"))
                    {
                        List<DBText> txt = null;
                        str[0] = item.TextString;
                        if (item.Rotation == 0)
                        {
                            txt = listText.Where(t => t.TextString.Contains(":")
                           && (Math.Floor(Math.Abs(t.Position.X-item.Position.X))==0)
                           && t.Rotation == item.Rotation && t.Height == item.Height &&
                           ((item.Position.Y - t.Position.Y) < 7 * item.Height) && (item.Position.Y >t.Position.Y) 
                           ).ToList();
                        }
                        else if (item.Rotation == Math.PI / 2)
                        {
                            txt = listText.Where(t => t.TextString.Contains(":")
                        && ((t.Position.X - item.Position.X) < 7 * item.Height) && (t.Position.X >item.Position.X)
                        && t.Rotation == item.Rotation && t.Height == item.Height &&
                       (Math.Floor(Math.Abs(item.Position.Y - t.Position.Y))==0)).ToList();
                        }
                        else
                        {
                            txt = listText.Where(t => t.TextString.Contains(":")
                        && ((t.Position.X - item.Position.X) / Math.Sin(item.Rotation)) < 7 * item.Height && ((t.Position.X - item.Position.X) / Math.Sin(item.Rotation)) >= 0
                        && t.Rotation == item.Rotation && t.Height == item.Height &&
                        ((item.Position.Y - t.Position.Y) / Math.Cos(item.Rotation)) < 7 * item.Height && ((item.Position.Y - t.Position.Y) / Math.Cos(item.Rotation)) >= 0
                        ).ToList();
                        }
                        if (txt.Count > 0)
                        {
                            for (int j = 0; j < txt.Count; j++)
                            {
                                switch (txt[j].TextString.Substring(0, 2))
                                {
                                    case "T:": str[1] = txt[j].TextString; break;
                                    case "M:": str[2] = txt[j].TextString; break;
                                    case "Q:": str[3] = txt[j].TextString; break;
                                    case "SI": str[4] = txt[j].TextString; break;
                                    case "MI": str[5] = txt[j].TextString; break;
                                }
                            }
                        }
                        str[6] = fileName;
                        listPartInfor.Add(str);
                        listText.Remove(item);
                    }
                    else continue;
                }
            }
            return listPartInfor;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

