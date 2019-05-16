using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Geometry;
using System.Text.RegularExpressions;
using System.IO;

namespace KFWH_CP_PlugIn
{
    public class KFWH_Elect_Cable
    {
        Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
        Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

        private readonly dynamic acadApp = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
        private readonly dynamic ThisDrawing = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
        //[CommandMethod("El_ExportOutCableTagList")]//20190114
        public void EL_导出cableList()
        {
            //将用户坐标系转换成世界坐标系
            this.acadApp.ZoomExtents();
            if (ed.CurrentUserCoordinateSystem != Matrix3d.Identity)
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
                ed.Regen();
            }
            const string TitleBlockName = "A3-PG"; const string reflinelayerName = "0";

            List<CableTag> listCableTags = new List<CableTag>();
            try
            {
                #region//读取数据
                using (Transaction trans = acDB.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(acDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                    if (bt.Has(TitleBlockName))
                    {
                        BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                        foreach (ObjectId id in btr)
                        {
                            var ent = id.GetObject(OpenMode.ForWrite) as Entity;

                            if (ent is BlockReference)
                            {
                                BlockReference blkref = ent as BlockReference;
                                if (blkref.Name == TitleBlockName)
                                {
                                    //获取名称
                                    string dwgNo = ""; string AltNo = "";
                                    var atts = blkref.AttributeCollection;
                                    if (atts != null)
                                    {
                                        foreach (ObjectId attid in atts)
                                        {
                                            AttributeReference attref = attid.GetObject(OpenMode.ForWrite) as AttributeReference;
                                            if (attref.Tag == "DWGNO") dwgNo = attref.TextString;
                                            if (attref.Tag == "ALT") AltNo = attref.TextString;
                                            attref.DowngradeOpen();
                                        }
                                    }

                                    double scale = blkref.ScaleFactors.X; if (scale<1) scale = scale * 25.4;

                                    //获取图框内的线
                                    PromptSelectionResult psr = ed.SelectWindow(blkref.GeometricExtents.MinPoint, blkref.GeometricExtents.MaxPoint,
                                        new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), new TypedValue((int)DxfCode.LayerName, reflinelayerName),
                                        new TypedValue((int)DxfCode.Color,201) }));
                                    if (psr.Status == PromptStatus.OK)
                                    {
                                        foreach (SelectedObject item in psr.Value)
                                        {
                                            CableTag ct = new CableTag(item.ObjectId, acDB, scale) { DwgNo = dwgNo, Rev = "R" + AltNo };
                                            if (ct.HasCableTagInfo) listCableTags.Add(ct);
                                        }
                                    }
                                    else
                                    {
                                        Circle c = new Circle();
                                        c.Center = new Point3d((blkref.GeometricExtents.MinPoint.X + blkref.GeometricExtents.MaxPoint.X) / 2, (blkref.GeometricExtents.MinPoint.Y + blkref.GeometricExtents.MaxPoint.Y) / 2, blkref.GeometricExtents.MinPoint.Z);
                                        c.Radius = blkref.GeometricExtents.MaxPoint.X - blkref.GeometricExtents.MinPoint.X;
                                        btr.AppendEntity(c);
                                        trans.AddNewlyCreatedDBObject(c, true);
                                        c.ColorIndex = 5;
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        ed.WriteMessage("找不到块为\"A3-PG\"的块！");
                    }
                    trans.Commit();
                }
                #endregion
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.Message);
            }

            if (listCableTags.Count > 0)
            {
                Excel.Application xlapp = new Excel.Application();
                try
                {
                    #region //导出到excek
                    Excel.Workbook wb = xlapp.Workbooks.Add();
                    Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                    ws.Range["A1:M1"].Value = new string[] { "DwgNo", "DwgRev", "CABLE TAG", "CABLE CODE", "CABLE LENGTH (m)",
                        "CABLE TYPE", "CABLE OD", "SIGNAL TYPE", "FROM EQUIP.","FROM EQUIP. LOCATION","TO EQUIP.","TO EQUIP. LOCATION","Remark" };
                    int k = 2;
                    foreach (var item in listCableTags)
                    {
                        if (item.CableTagNumbers != null)
                        {
                            for (int i = 0; i < item.CableTagNumbers.Length; i++)
                            {
                                ws.Range["A" + k].Value = item.DwgNo;
                                ws.Range["B" + k].Value = item.Rev;
                                ws.Range["C" + k].Value = item.CableTagNumbers[i];
                                ws.Range["F" + k].Value = item.CableTagSize;
                                ws.Range["I" + k].Value = item.FromEquimentDes[0].Replace("%%U","").Replace("%%D"," Degre e").Replace("%%C", "diameter");
                                ws.Range["J" + k].Value = item.FromEquimentDes[1].Replace("%%U", "").Replace("%%D", " Degre e").Replace("%%C", "diameter");
                                ws.Range["K" + k].Value = item.ToEquimentDes[0].Replace("%%U", "").Replace("%%D", " Degre e").Replace("%%C", "diameter");
                                ws.Range["L" + k].Value = item.ToEquimentDes[1].Replace("%%U", "").Replace("%%D", " Degre e").Replace("%%C", "diameter");
                                k++;
                            }
                        }

                    }
                    ws.Range["A1:M1"].EntireColumn.AutoFit();
                    ws.Range["A1:M1"].EntireColumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlapp.Visible = true;
                    xlapp.WindowState = Excel.XlWindowState.xlMaximized;
                    #endregion
                }
                catch (System.Exception ex)
                {
                    xlapp.Quit();
                    ed.WriteMessage(ex.Message);
                }
            }
        }
    }
}

public class CableTag
{
    public string[] CableTagNumbers { get; set; }
    public string CableTagSize { get; set; }
    public string[] FromEquimentDes { get; set; }
    public string[] ToEquimentDes { get; set; }
    public string DwgNo { get; set; }
    public string Rev { get; set; }
    public const string CableTagBlkName = "";
    public bool HasCableTagInfo { get; set; }
    public CableTag(ObjectId plineObjID, Database acDb, double TtitleBlockscale)
    {
        using (Transaction trans = acDb.TransactionManager.StartTransaction())
        {
            var pline = trans.GetObject(plineObjID, OpenMode.ForWrite) as Polyline;
            var sp = pline.StartPoint;
            var ep = pline.EndPoint;
            var mp = pline.GetPoint3dAt(1);
            Editor ed = acadApp.DocumentManager.GetDocument(acDb).Editor;
            //选择文字--->起始设备
            #region
            PromptSelectionResult psr = ed.SelectCrossingWindow(sp, new Point3d(sp.X + (4 * TtitleBlockscale), sp.Y + (10 * TtitleBlockscale), sp.Z),
                new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Text") }));
            if (psr.Status == PromptStatus.OK)
            {
                List<DBText> listTxt = new List<DBText>();
                foreach (SelectedObject item in psr.Value)
                {
                    var txt = trans.GetObject(item.ObjectId, OpenMode.ForWrite) as DBText;
                    txt.ColorIndex = 171;
                    listTxt.Add(txt);
                }
                var temp = listTxt.OrderBy(c => c.Position.Y).ToList();
                if (temp.Count > 0)
                {
                    this.FromEquimentDes = new string[2];
                    this.FromEquimentDes[1] = temp[0].TextString;
                    for (int i = temp.Count - 1; i > 0; i--)
                    {
                        this.FromEquimentDes[0] += temp[i].TextString + ",";
                    }
                    this.FromEquimentDes[0] = this.FromEquimentDes[0].Substring(0, this.FromEquimentDes[0].Length - 1);
                }
                this.HasCableTagInfo = true;
            }
            else this.HasCableTagInfo = false;
            #endregion
            //选择文字--->终止设备
            #region
            psr = ed.SelectCrossingWindow(ep, new Point3d(ep.X + (4 * TtitleBlockscale), ep.Y + (10 * TtitleBlockscale), ep.Z),
                    new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Text") }));
            if (psr.Status == PromptStatus.OK)
            {
                List<DBText> listTxt = new List<DBText>();
                foreach (SelectedObject item in psr.Value)
                {
                    var txt = trans.GetObject(item.ObjectId, OpenMode.ForWrite) as DBText;
                    txt.ColorIndex = 171;
                    listTxt.Add(txt);
                }
                var temp = listTxt.OrderBy(c => c.Position.Y).ToList();
                if (temp.Count > 0)
                {
                    this.ToEquimentDes = new string[2];
                    this.ToEquimentDes[1] = temp[0].TextString;
                    for (int i = temp.Count - 1; i > 0; i--)
                    {
                        this.ToEquimentDes[0] += temp[i].TextString + ",";
                    }
                    this.ToEquimentDes[0] = this.ToEquimentDes[0].Substring(0, this.ToEquimentDes[0].Length - 1);
                }
                this.HasCableTagInfo = true;
            }
            else this.HasCableTagInfo = false;
            #endregion
            //选择文字或图块--->电缆号码和规格
            #region
            psr = ed.SelectCrossingWindow(mp, new Point3d(mp.X + (4 * TtitleBlockscale), mp.Y + (10 * TtitleBlockscale), mp.Z),
                    new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Text") }));
            if (psr.Status == PromptStatus.OK)
            {
                List<DBText> listTxt = new List<DBText>();
                foreach (SelectedObject item in psr.Value)
                {
                    var txt = trans.GetObject(item.ObjectId, OpenMode.ForWrite) as DBText;
                    txt.ColorIndex = 171;
                    listTxt.Add(txt);
                }
                var temp = listTxt.OrderBy(c => c.Position.Y).ToList();
                if (temp.Count > 0)
                {
                    this.CableTagSize = temp[0].TextString;
                    this.CableTagNumbers = new string[temp.Count - 1];
                    for (int i = 1; i < temp.Count; i++)
                    {
                        this.CableTagNumbers[i - 1] += temp[i].TextString;
                    }
                }
                this.HasCableTagInfo = true;
            }
            else
            {
                this.HasCableTagInfo = false;
                psr = ed.SelectCrossingWindow(mp, new Point3d(mp.X + (4 * TtitleBlockscale), mp.Y + (5 * TtitleBlockscale), mp.Z),
                                        new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Insert") }));
                if (psr.Status == PromptStatus.OK && psr.Value.Count == 1)
                {
                    List<string> listTxt = new List<string>();
                    var blk = trans.GetObject(psr.Value[0].ObjectId, OpenMode.ForWrite) as BlockReference;
                    blk.ColorIndex = 171;
                    if (blk.AttributeCollection != null)
                    {
                        foreach (ObjectId item in blk.AttributeCollection)
                        {
                            AttributeReference attref = item.GetObject(OpenMode.ForWrite) as AttributeReference;
                            if (attref.Tag.ToUpper().Contains("CABLENUMBER") || attref.Tag.ToUpper() == "CABLETYPE") listTxt.Add(attref.TextString);
                            attref.DowngradeOpen();
                        }
                        var temp = listTxt.Where(c => c.Contains("(") && c.Contains(")")).ToList();
                        if (temp.Count > 0)
                        {
                            listTxt.Remove(temp[0]);
                            this.CableTagSize = temp[0];
                            this.CableTagNumbers = new string[listTxt.Count];
                            for (int i = 0; i < listTxt.Count; i++)
                            {
                                this.CableTagNumbers[i] = listTxt[i];
                            }

                        }
                    }
                    this.HasCableTagInfo = true;
                }
                else this.HasCableTagInfo = false;
            }
            #endregion
            if (this.HasCableTagInfo)
            {
                pline.ColorIndex = 171;
                trans.Commit();
            }
            else trans.Abort();

        }
    }
}
