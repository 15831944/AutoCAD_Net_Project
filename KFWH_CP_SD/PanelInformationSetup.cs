using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using KFWH_CMD;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Reflection;

namespace KFWH_CP_SD
{
    public partial class PanelInformationSetup : Form
    {
        public PanelInfor CurPanel { get; set; }

        public PanelInformationSetup()
        {
            InitializeComponent();
        }

        private void comboBox1_DropDown(object sender, EventArgs e)
        {
            this.comboBox1.Items.Clear();
            this.comboBox1.Items.AddRange(MySQLHelper.GetAllProjectName().ToArray());
        }

        private void comboBox2_DropDown(object sender, EventArgs e)
        {
            this.comboBox2.Items.Clear();
            if (this.comboBox1.SelectedIndex == -1)
            {
                this.comboBox1.Focus();
            }
            else this.comboBox2.Items.AddRange(MySQLHelper.GetCurProjectBlks(this.comboBox1.Text).ToArray());
        }

        private void comboBox3_DropDown(object sender, EventArgs e)
        {
            this.comboBox3.Items.Clear();
            if (this.comboBox1.SelectedIndex == -1 || this.comboBox2.SelectedIndex == -1)
            {
                if (this.comboBox1.SelectedIndex == -1) this.comboBox1.Focus(); else this.comboBox2.Focus();
            }
            else this.comboBox3.Items.AddRange(MySQLHelper.GetCurProjectBlkPanels(this.comboBox1.Text, this.comboBox2.Text).ToArray());
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.CurPanel = MySQLHelper.GetPanelTitleBlockInfo(this.comboBox1.Text, this.comboBox2.Text, this.comboBox3.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!this.textBox1.Text.IsNullOrWhiteSpace() && this.textBox4.Text.IsNumeric())
            {
                this.CurPanel.DrawingNo = this.textBox1.Text;
                this.CurPanel.Date = String.Format("{0:00}", DateTime.Now.Day) + "/" + String.Format("{0:00}", DateTime.Now.Month) + "/" + String.Format("{0:0000}", DateTime.Now.Year);
                this.CurPanel.Alt = this.textBox4.Text;
                using (DocumentLock doclock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    var db = HostApplicationServices.WorkingDatabase;
                    List<string> properties = typeof(PanelInfor).GetProperties().Select(p => p.Name).ToList();
                    Dictionary<string, string> dic = new Dictionary<string, string>();
                    foreach (var item in properties) dic.Add(item, typeof(PanelInfor).GetProperty(item).GetValue(this.CurPanel).ToString());
                    if (db.HasSummaryInfo())
                    {
                        var info = new DatabaseSummaryInfoBuilder();
                        foreach (KeyValuePair<string, string> item in dic) info.CustomPropertyTable.Add(item.Key, item.Value);
                        db.SummaryInfo = info.ToDatabaseSummaryInfo();
                    }
                    else AddCustomInfo(dic);
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Regen();
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Panel Inforamtion Update Complate!");
                }
            }
            else
            {
                if (this.textBox1.Text.IsNullOrWhiteSpace()) { this.textBox1.Focus(); Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("please input correct drawing Number"); }
                if (!this.textBox4.Text.IsNumeric()) { this.textBox4.Focus(); Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Please Input Hull Drawing Alteration"); }
            }

        }

        private void AddCustomInfo(Dictionary<string,string> dic)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            if (db.HasSummaryInfo()) return;
            var info = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
            foreach (KeyValuePair<string,string> item in dic) info.CustomPropertyTable.Add(item.Key, item.Value);
            db.SummaryInfo = info.ToDatabaseSummaryInfo();
        }
    }
}
