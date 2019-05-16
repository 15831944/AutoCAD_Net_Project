using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Windows;
using System;

using System.Linq;
using System.Text;
using System.Windows;
//using System.Windows.Media.Imaging;
using System.Data.OleDb;
using KFWH_CP_PlugIn;
using System.Resources;
using AutoCAD_Net_Project;
using System.Data.SqlClient;

using KFWH_CMD;


// This line is not mandatory, but improves loading performances
namespace AutoCAD_Net_Project
{
    public partial class MTO
    {
        [KFWH_CMD.myAutoCADcmd("MTO 批量取零件，编号", "MTO_CreateParts.bmp")]
        [CommandMethod("MTO_CreateParts", CommandFlags.UsePickSet)]
        public void MTO取Parts() // This method can have any name
        {

            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSNAPCOORD", 0);
            if (Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CECOLOR").ToString() != "BYLAYER") Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CECOLOR", "BYLAYER");
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //将用户坐标系转换成世界坐标系
            if (ed.CurrentUserCoordinateSystem != Matrix3d.Identity)
            {
                ed.CurrentUserCoordinateSystem = Matrix3d.Identity;
                ed.Regen();
            }
            string viewName = ed.GetString("Please input drawing name, such as: BB/LB3500/FR27/MDK").StringResult;
            string partType = "";
            var psrKey = ed.GetKeywords("Please a Type as parts Type: ", new string[] { "HPlate", "TWeb", "BKT", "Hold" });
            if (psrKey.Status == PromptStatus.OK)
            {
                partType = psrKey.StringResult;
            }
            DBObjectCollection ptC = new DBObjectCollection();//parts的边界集合
            List<double[]> listp = new List<double[]>();
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.MessageForAdding = "Please select all text in Blk* layer";
            Dictionary<string, ObjectId[]> dicParts = new Dictionary<string, ObjectId[]>();
            PromptSelectionResult psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start,"Text"),
                new TypedValue((int)DxfCode.LayerName,"Blk*")}));
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet ss = psr.Value;
                using (Transaction trans3 = acDB.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTableRecord btr = trans3.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        foreach (SelectedObject item in ss)
                        {
                            var obj = trans3.GetObject(item.ObjectId, OpenMode.ForRead);
                            ptC.Add(obj);
                        }
                        int x = 0;
                        //获取块表
                        BlockTable cutBT = trans3.GetObject(acDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                        if (cutBT.Has("PartsInfor")) x = 1;
                        if (x != 1)
                        {
                            List<string> listDlls = new List<string>();
                            string sqlcommand = $"select * from tb_PluginUsedFiles where Class_Name='MTO_Dynamic_Block' and FileName='MTO_Part_DynamicBlock.dwg'";
                            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
                            if (sdr.HasRows)//判断行不为空
                            {
                                while (sdr.Read())//循环读取数据，知道无数据为止
                                {
                                    listDlls.Add(sdr["Path"].ToString() + "\\" + sdr["FileName"].ToString());
                                }
                            }
                            string partBlock = listDlls[0];
                            acDB.ImportBlocksFromDwg(partBlock);
                        }
                        trans3.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        trans3.Abort();
                        ed.WriteMessage(ex.Message);
                    }
                    ss.Dispose();
                }
                DBObjectCollection partCollection = new DBObjectCollection();
                using (Transaction trans = acDB.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        int i = 0;
                        foreach (DBObject item in ptC)
                        {
                            i++;
                            var text = item as DBText;
                            var pnts = text.GeometricExtents;
                            Point3d pnt3d = new Point3d(0.5 * pnts.MinPoint.X + 0.5 * pnts.MaxPoint.X, 0.5 * pnts.MinPoint.Y + 0.5 * pnts.MaxPoint.Y, 0.5 * pnts.MinPoint.Z + 0.5 * pnts.MaxPoint.Z);
                            DBObjectCollection dbc;
                            try
                            {
                                //ed.Regen();
                                dbc = ed.TraceBoundary(pnt3d, true); //生成边界
                            }
                            catch (System.Exception ex)
                            {
                                dbc = new DBObjectCollection();
                                ed.WriteMessage(ex.Message);
                            }
                            if (dbc.Count == 1)//零件只能出现一个
                            {
                                using (Transaction trans1 = acDB.TransactionManager.StartTransaction())
                                {
                                    BlockTableRecord btr1 = trans1.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                    foreach (DBObject item1 in dbc)
                                    {
                                        var ent = item1 as Entity;
                                        ent.Layer = text.Layer;
                                        if (ent is Region)
                                        {
                                            if (dicParts.ContainsKey(i + "|" + (ent as Region).Area.ToString()))
                                            {
                                                break;
                                            }
                                            else
                                            {
                                                dicParts.Add(i + "|" + (ent as Region).Area.ToString(), new ObjectId[] { btr1.AppendEntity(ent), text.ObjectId });
                                                trans1.AddNewlyCreatedDBObject(ent, true);
                                            }
                                        }
                                        else
                                        {
                                            if (dicParts.ContainsKey(i + "|" + (ent as Curve).Area.ToString()))
                                            {
                                                break;
                                            }
                                            else
                                            {
                                                dicParts.Add(i + "|" + (ent as Curve).Area.ToString(), new ObjectId[] { btr1.AppendEntity(ent), text.ObjectId });
                                                trans1.AddNewlyCreatedDBObject(ent, true);
                                            }
                                        }
                                    }
                                    trans1.Commit();
                                }//bo做出ployline
                            }
                            else listp.Add(new double[] { pnt3d.X, pnt3d.Y, pnt3d.Z });
                        }
                        trans.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        trans.Abort();
                        ed.WriteMessage(ex.Message);
                    }
                }
                using (Transaction trans2 = acDB.TransactionManager.StartTransaction())
                {
                    try
                    {
                        BlockTableRecord btr2 = trans2.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        int PL = 1, bk = 1, hold = 1, web = 1;
                        foreach (KeyValuePair<string, ObjectId[]> keyValues in dicParts)
                        {
                            ObjectId blkrefID = new ObjectId();
                            var entBos = keyValues.Value[0].GetObject(OpenMode.ForRead);
                            Entity entBoundary;
                            if (entBos is Curve) entBoundary = entBos as Curve; else entBoundary = entBos as Region;

                            var thk_Grade = keyValues.Value[1].GetObject(OpenMode.ForRead) as DBText;
                            var pnts = thk_Grade.GeometricExtents;
                            Point3d centerPnt = new Point3d(pnts.MaxPoint.X * 0.5 + pnts.MinPoint.X * 0.5,
                              pnts.MaxPoint.Y * 0.5 + pnts.MinPoint.Y * 0.5,
                              pnts.MaxPoint.Z * 0.5 + pnts.MinPoint.Z * 0.5);
                            blkrefID = MTO.InsertBlockref(acDB.CurrentSpaceId, entBoundary.Layer, "PartsInfor", centerPnt, new Scale3d(thk_Grade.Height / 9,
                               thk_Grade.Height / 9,
                               thk_Grade.Height / 9), thk_Grade.Rotation);
                            BlockReference partBlockref = blkrefID.GetObject(OpenMode.ForWrite) as BlockReference;
                            if (partBlockref != null)
                            {
                                foreach (ObjectId item in partBlockref.AttributeCollection)
                                {
                                    AttributeReference attref = item.GetObject(OpenMode.ForWrite) as AttributeReference;
                                    switch (attref.Tag)
                                    {
                                        case "THK_WITHGR":
                                            //判断厚度为数字，拉线
                                            var thicknessValue = thk_Grade.TextString.Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                            double thk = 0;
                                            if (thicknessValue.Length > 0)
                                            {
                                                if (!double.TryParse(thicknessValue[0].Trim(), out thk) && !thicknessValue[0].Contains("\""))
                                                {
                                                    listp.Add(new double[] { partBlockref.Position.X, partBlockref.Position.Y, partBlockref.Position.Z });
                                                    attref.TextString = thk_Grade.TextString;
                                                    attref.Layer = partBlockref.Layer;
                                                }
                                                else
                                                {
                                                    attref.TextString = thk_Grade.TextString;
                                                    attref.Layer = partBlockref.Layer;
                                                }
                                            }
                                            break;
                                        case "PIECENAME":
                                            switch (partType)
                                            {
                                                //{ "HPlate","TWeb","BKT","Hold" });:
                                                case "HPlate":
                                                    attref.TextString = thk_Grade.Layer + "/" + viewName + "-H" + PL;
                                                    PL++;
                                                    attref.Layer = partBlockref.Layer;
                                                    break;
                                                case "TWeb":
                                                    attref.TextString = thk_Grade.Layer + "/" + viewName + "-TW" + web;
                                                    web++;
                                                    attref.Layer = partBlockref.Layer;
                                                    break;
                                                case "BKT":
                                                    attref.TextString = thk_Grade.Layer + "/" + viewName + "-BK" + bk;
                                                    bk++;
                                                    attref.Layer = partBlockref.Layer;
                                                    break;
                                                case "Hold":
                                                    attref.TextString = thk_Grade.Layer + "/" + viewName + "-PL" + hold;
                                                    hold++;
                                                    attref.Layer = partBlockref.Layer;
                                                    break;
                                            }
                                            break;
                                        case "QTY":
                                            attref.TextString = "N=1";
                                            attref.Layer = partBlockref.Layer;
                                            break;
                                        case "ASS":
                                            attref.TextString = partType;
                                            attref.Layer = partBlockref.Layer;
                                            break;
                                        case "AREA":
                                            attref.Layer = partBlockref.Layer;
                                            attref.TextString = keyValues.Key.Split(new char[] { '|' })[1];
                                            break;
                                    }
                                    attref.DowngradeOpen();
                                }
                            }
                            //DotNetARX.EntTools.Erase(thk_Grade.ObjectId);
                        }
                        trans2.Commit();
                        if (listp.Count > 0)
                        {
                            foreach (var item in listp)
                            {
                                Line l = new Line(new Point3d(item[0], item[1], item[2]), Point3d.Origin);
                                acDB.AddToModelSpace(l);
                            }
                        }
                        sw.Stop();
                        if (dicParts.Count > 0)
                        {
                            ed.WriteMessage("Total " + (dicParts.Count) + " parts create completed!" + System.Environment.NewLine);
                            ed.WriteMessage("Total " + sw.ElapsedMilliseconds / 1000000 + " seconds cost");
                        }
                        else
                        {
                            ed.WriteMessage("Please copy your part to an new AutoCAD file, and run command again!");
                        }

                    }
                    catch (System.Exception)
                    {
                        trans2.Abort();
                        ed.WriteMessage("some error were been accour!");
                    }
                }
            }
        }

        [KFWH_CMD.myAutoCADcmd("MTO 板材与套料程序", "MTO_AutoNesting.bmp")]
        [CommandMethod("MTO_AutoNesting")]
        public void MTO板材预套料()
        {
            Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("OSNAPCOORD", 0);
            if (Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem != Autodesk.AutoCAD.Geometry.Matrix3d.Identity)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem = Autodesk.AutoCAD.Geometry.Matrix3d.Identity;
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Regen();
            }
            Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            double scale = 1;
            if (acDB.Lunits == 4) scale = 25.4;// 1==>Scientific;2==>Decimal;3==>Engineering;4==> Architectural;5==> Fractional

            if (Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CECOLOR").ToString() != "BYLAYER") Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("CECOLOR", "BYLAYER");
            PromptSelectionOptions pso = new PromptSelectionOptions();
            PromptSelectionResult psr = null;
            pso.MessageForAdding = "Please select one All Parts(Only Circle/Polyline/Region) Can be select!：";
            pso.AllowDuplicates = false; pso.SingleOnly = false;
            //
            List<NCParts> listNcPart = new List<NCParts>();
            psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "Circle,LWPOLYLINE,Region"),
                new TypedValue((int)DxfCode.LayerName, "Blk*")
            }));
            if (psr.Status == PromptStatus.OK)
            {
                using (Transaction trans = acDB.TransactionManager.StartTransaction())
                {
                    #region// 收集零件信息
                    BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    List<DBObject> listcur = new List<DBObject>();
                    foreach (SelectedObject item in psr.Value)
                    {
                        listcur.Add(trans.GetObject(item.ObjectId, OpenMode.ForRead));
                    }
                    foreach (var item in listcur)
                    {
                        Line l1 = null; Line l2 = null; NCParts ncpart = null;
                        if (item is Curve)
                        {
                            Curve cur = item as Curve;
                            ncpart = new NCParts(acDB, cur);
                            if (ncpart.HasPartInfo)
                            {
                                if (ncpart.PartInfoQty == 1)
                                {
                                    listNcPart.Add(ncpart);
                                }
                                else
                                {
                                    l1 = new Line();
                                    l1.StartPoint = Point3d.Origin;
                                    l1.EndPoint = cur.GetCpnt();
                                    l1.Layer = "0";
                                    l1.ColorIndex = 1;
                                }
                            }
                            else
                            {
                                l2 = new Line();
                                l2.StartPoint = Point3d.Origin;
                                l2.EndPoint = cur.GetCpnt();
                                l2.Layer = "0";
                                l2.ColorIndex = 2;
                            }
                        }
                        else
                        {
                            Region re = item as Region;
                            ncpart = new NCParts(acDB, re);
                            if (ncpart.HasPartInfo)
                            {
                                if (ncpart.PartInfoQty == 1)
                                {
                                    listNcPart.Add(ncpart);
                                }
                                else
                                {
                                    l1 = new Line();
                                    l1.StartPoint = Point3d.Origin;
                                    l1.EndPoint = re.GetCpnt();
                                    l1.Layer = "0";
                                    l1.ColorIndex = 1;
                                }
                            }
                            else
                            {
                                l2 = new Line();
                                l2.StartPoint = Point3d.Origin;
                                l2.EndPoint = re.GetCpnt();
                                l2.Layer = "0";
                                l2.ColorIndex = 2;
                            }
                        }
                        if (l1 != null) { acDB.AddToModelSpace(l1); }
                        if (l2 != null) { acDB.AddToModelSpace(l2); }
                    }
                    #endregion

                    ed.WriteMessage("Read Part information Complete , " + $"red color line mean Name not unique,yellow color means Can't Find Part Name Label!");

                    if (listNcPart.Count > 0)
                    {
                        Dictionary<string, List<NCParts>> dicNcParts = new Dictionary<string, List<NCParts>>();
                        var sortParts = listNcPart.OrderBy(c => c.BlkName).ThenBy(c => c.THK).ThenBy(c => c.Grade).ThenBy(c => c.PartType).ThenBy(c => c.Area).ToList();
                        PromptPointResult ppr = ed.GetPoint("Please pick a point to Put Nesting Result:");
                        if (ppr.Status == PromptStatus.OK)
                        {
                            #region//插入块到数据库
                            if (scale == 1)
                            {
                                List<string> listDlls = new List<string>();
                                string sqlcommand = $"select * from tb_PluginUsedFiles where Class_Name='MTO_Dynamic_Block' and FileName='MTO_Material_Plate.dwg'";
                                SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
                                if (sdr.HasRows)//判断行不为空
                                {
                                    while (sdr.Read())//循环读取数据，知道无数据为止
                                    {
                                        listDlls.Add(sdr["Path"].ToString() + "\\" + sdr["FileName"].ToString());
                                    }
                                }
                                string partBlock = listDlls[0];
                                InsertBlockFromExistingDrawing("Material_Plate", partBlock);
                            }
                            else
                            {

                                List<string> listDlls = new List<string>();
                                string sqlcommand = $"select * from tb_PluginUsedFiles where Class_Name='MTO_Dynamic_Block' and FileName='MTO_Material_Plate_Inch.dwg'";
                                SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
                                if (sdr.HasRows)//判断行不为空
                                {
                                    while (sdr.Read())//循环读取数据，知道无数据为止
                                    {
                                        listDlls.Add(sdr["Path"].ToString() + "\\" + sdr["FileName"].ToString());
                                    }
                                }
                                string partBlock = listDlls[0];
                                InsertBlockFromExistingDrawing("Material_Plate_Inch", partBlock);
                            }
                            #endregion
                            //分类
                            #region // 零件分类
                            foreach (var item in sortParts)
                            {
                                if (dicNcParts.ContainsKey(item.BlkName + "|" + item.THK + "|" + item.Grade + "|" + item.PartType))
                                {
                                    var listTemp = dicNcParts[item.BlkName + "|" + item.THK + "|" + item.Grade + "|" + item.PartType];
                                    listTemp.Add(item);
                                    dicNcParts[item.BlkName + "|" + item.THK + "|" + item.Grade + "|" + item.PartType] = listTemp.OrderBy(c => c.Area).ThenBy(c => c.Length).ThenBy(c => c.Width).ToList();
                                }
                                else
                                {
                                    var listTemp = new List<NCParts>();
                                    listTemp.Add(item);
                                    dicNcParts[item.BlkName + "|" + item.THK + "|" + item.Grade + "|" + item.PartType] = listTemp;
                                }
                            }
                            //套入
                            #endregion 
                            var insertPnt = ppr.Value;
                            for (int i = 0; i < dicNcParts.Count; i++)
                            {
                                Point3d insPnt = new Point3d(insertPnt.X + (i * 20000 / scale), insertPnt.Y, insertPnt.Z);
                                KeyValuePair<string, List<NCParts>> keyvalue = new KeyValuePair<string, List<NCParts>>(dicNcParts.Keys.ToList()[i], dicNcParts.Values.ToList()[i]);
                                var listParts = keyvalue.Value.ToList();
                                while (listParts.Count > 0)
                                {
                                    using (Transaction trans1 = acDB.TransactionManager.StartTransaction())
                                    {
                                        RawMaterial ram = new RawMaterial(listParts[0], acDB, insPnt, 9144, 2438, scale);
                                        ram.NestedNCPartsToRaw(listParts);
                                        insPnt = new Point3d(insPnt.X, insPnt.Y + 8000 / scale, insPnt.Z);
                                        trans1.Commit();
                                    }
                                }
                            }
                        }
                        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Regen();
                    }
                    trans.Commit();
                }
            }
        }

        private ObjectId InsertBlockFromExistingDrawing(string blkName, string dwgFileName)//"Material_Plate",@"\\10.19.80.8\cnkfewhfls01\Users\sheng.nan\AutoProfileNesting\AutoProfileNesting\Filing of Block for Material Plate.dwg"
        {
            ObjectId btrid = ObjectId.Null;
            Database curDb = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = curDb.TransactionManager.StartTransaction())
            {
                BlockTable curBT = trans.GetObject(curDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                
                if (!curBT.Has(blkName))
                {
                    using (Database db = new Database(false, true))
                    {
                        db.ReadDwgFile(dwgFileName, System.IO.FileShare.Read, true, null);
                        db.CloseInput(true);
                        btrid = curDb.Insert(blkName, db, false);
                    }
                }
                trans.Commit();
            }
            return btrid;
        }
        public static ObjectId InsertBlockref(ObjectId spacdId, string layer, string blkName, Point3d position, Scale3d scale, double roAngle)
        {
            ObjectId blkrefID;
            Database db = spacdId.Database;
            BlockTable bt = db.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;
            if (!bt.Has(blkName)) return ObjectId.Null;
            BlockTableRecord space = spacdId.GetObject(OpenMode.ForWrite) as BlockTableRecord;
            blkrefID = bt[blkName];
            BlockTableRecord record = blkrefID.GetObject(OpenMode.ForRead) as BlockTableRecord;
            BlockReference br = new BlockReference(position, blkrefID);
            br.ScaleFactors = scale;
            br.Layer = layer;
            br.Rotation = roAngle;
            space.AppendEntity(br);
            if (record.HasAttributeDefinitions)
            {
                foreach (ObjectId id in record)
                {
                    var attdef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attdef != null)
                    {
                        AttributeReference ar = new AttributeReference();
                        ar.SetAttributeFromBlock(attdef, br.BlockTransform);
                        ar.Position = attdef.Position.TransformBy(br.BlockTransform);
                        ar.Rotation = roAngle;
                        ar.AdjustAlignment(db);
                        ar.Layer = layer;
                        br.AttributeCollection.AppendAttribute(ar);
                        db.TransactionManager.AddNewlyCreatedDBObject(ar, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            return br.ObjectId;
        }

    }
    public class NCParts
    {
        public int Quantity { get; set; }
        public string PartName { get; set; }
        public string PartType { get; set; }
        public int PartInfoQty { get; set; }
        public bool HasPartInfo { get; set; }
        public double THK { get; set; }
        public string Grade { get; set; }
        public double Area { get; set; }
        public Database AcDb { get; set; }
        public ObjectId ShapeObjectID { get; set; }
        public ObjectId LabelObjectID { get; set; }
        public string BlkName { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public double OptimizationAngle { get; set; }
        public Point3d centerPnt { get; set; }
        public NCParts(Database acDb, Curve cur)
        {
            this.AcDb = acDb;
            this.BlkName = cur.Layer;
            this.ShapeObjectID = cur.ObjectId;
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                PromptSelectionResult psr = null;
                SelectionFilter sf = new SelectionFilter(new TypedValue[] {
                    new TypedValue((int)DxfCode.LayerName,"Blk*"),
                    new TypedValue((int)DxfCode.BlockName,"PartsInfor") }
                );
                if (cur is Polyline)
                {
                    var pline = cur as Polyline;
                    psr = acDb.GetEditor().SelectCrossingPolygon(pline.GetAllPoints(), sf);
                }
                if (psr.Status != PromptStatus.OK || psr == null) psr = acDb.GetEditor().SelectCrossingWindow(cur.GeometricExtents.MaxPoint, cur.GeometricExtents.MinPoint, sf);
                if (psr.Status == PromptStatus.OK)
                {
                    BlockReference blk = null;
                    this.HasPartInfo = true;
                    this.PartInfoQty = psr.Value.Count;
                    if (this.PartInfoQty == 1)
                    {
                        #region//选择零件内部信息
                        this.LabelObjectID = psr.Value[0].ObjectId;
                        blk = trans.GetObject(this.LabelObjectID, OpenMode.ForRead) as BlockReference;
                        if (blk != null)
                        {
                            foreach (ObjectId item in blk.AttributeCollection)
                            {
                                AttributeReference attref = item.GetObject(OpenMode.ForWrite) as AttributeReference;
                                switch (attref.Tag)
                                {
                                    case "THK_WITHGR":
                                        var thicknessValue = attref.TextString.Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);

                                        if (thicknessValue.Length > 0)
                                        {
                                            if (thicknessValue[0].Contains("\""))
                                            {
                                                var arr = attref.TextString.Replace(thicknessValue[0], MyTools.InchThk2mm(thicknessValue[0].Trim()).ToString()).Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                                this.THK = double.Parse(arr[0]);
                                                this.Grade = arr[1];
                                            }
                                            else
                                            {
                                                this.THK = double.Parse(thicknessValue[0]);
                                                this.Grade = thicknessValue[1];
                                            }
                                        }

                                        break;
                                    case "PIECENAME":
                                        this.PartName = attref.TextString;
                                        break;
                                    case "QTY":
                                        this.Quantity = int.Parse(attref.TextString.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries)[1]);
                                        break;
                                    case "ASS":
                                        this.PartType = attref.TextString;
                                        break;
                                    case "AREA":
                                        this.Area = double.Parse(attref.TextString);
                                        break;
                                }
                            }
                        }
                        #endregion

                        #region //找出零件的最佳角度
                        if (cur is Polyline)
                        {
                            bool flag = false;
                            Point3dCollection pnts = new Point3dCollection();
                            List<double[]> pntDis = new List<double[]>();
                            for (int i = 0; i < (cur as Polyline).NumberOfVertices; i++)
                            {
                                var curPnt = (cur as Polyline).GetPoint3dAt(i);
                                if (i != (cur as Polyline).NumberOfVertices - 1)
                                {
                                    var nextPnt = (cur as Polyline).GetPoint3dAt(i + 1);
                                    pntDis.Add(new double[] { curPnt.DistanceTo(nextPnt), curPnt.X, curPnt.Y, nextPnt.X, nextPnt.Y });
                                }
                            }
                            this.centerPnt = new Point3d(
                                (cur.GeometricExtents.MaxPoint.X + cur.GeometricExtents.MinPoint.X) * 0.5,
                                 (cur.GeometricExtents.MaxPoint.Y + cur.GeometricExtents.MinPoint.Y) * 0.5,
                                  (cur.GeometricExtents.MaxPoint.Z + cur.GeometricExtents.MinPoint.Z) * 0.5
                                );

                            double l = cur.GeometricExtents.MaxPoint.X - cur.GeometricExtents.MinPoint.X;
                            double w = cur.GeometricExtents.MaxPoint.Y - cur.GeometricExtents.MinPoint.Y;

                            List<double[]> maxDis = pntDis.Where(p => p[0] == pntDis.Max(c => c[0])).ToList();
                            if (maxDis.Count(c => c[1] == c[3] || c[2] == c[4]) > 0)
                            {
                                if (l > w)
                                {
                                    this.Length = l; this.Width = w;
                                    this.OptimizationAngle = 0;
                                }
                                else
                                {
                                    this.Length = w; this.Width = l;
                                    this.OptimizationAngle = Math.PI * 0.5;
                                }
                            }
                            else
                            {
                                var detlaX = Math.Abs(maxDis.First()[1] - maxDis.First()[3]);
                                var detlaY = Math.Abs(maxDis.First()[2] - maxDis.First()[4]);
                                this.OptimizationAngle = Math.Atan(detlaY / detlaX);
                                using (Transaction trans1 = acDb.TransactionManager.StartTransaction())
                                {
                                    MyTools.MyRotate(cur, this.centerPnt, Math.PI * 2 - this.OptimizationAngle);
                                    this.Length = cur.GeometricExtents.MaxPoint.X - cur.GeometricExtents.MinPoint.X;
                                    this.Width = cur.GeometricExtents.MaxPoint.Y - cur.GeometricExtents.MinPoint.Y;
                                }
                            }
                        }
                        #endregion

                    }
                }
            }
        }
        public NCParts(Database acDb, Region cur)
        {
            this.AcDb = acDb;
            this.BlkName = cur.Layer;
            this.ShapeObjectID = cur.ObjectId;
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                PromptSelectionResult psr = null;
                SelectionFilter sf = new SelectionFilter(new TypedValue[] {
                    new TypedValue((int)DxfCode.LayerName,"Blk*"),
                    new TypedValue((int)DxfCode.BlockName,"PartsInfor") }
                );
                psr = acDb.GetEditor().SelectCrossingWindow(cur.GeometricExtents.MaxPoint, cur.GeometricExtents.MinPoint, sf);
                if (psr.Status == PromptStatus.OK)
                {
                    BlockReference blk = null;
                    this.HasPartInfo = true;
                    this.PartInfoQty = psr.Value.Count;
                    if (this.PartInfoQty == 1)
                    {
                        #region//选择零件内部信息
                        this.LabelObjectID = psr.Value[0].ObjectId;
                        blk = trans.GetObject(this.LabelObjectID, OpenMode.ForRead) as BlockReference;
                        if (blk != null)
                        {
                            foreach (ObjectId item in blk.AttributeCollection)
                            {
                                AttributeReference attref = item.GetObject(OpenMode.ForWrite) as AttributeReference;
                                switch (attref.Tag)
                                {
                                    case "THK_WITHGR":
                                        var thicknessValue = attref.TextString.Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);

                                        if (thicknessValue.Length > 0)
                                        {
                                            if (thicknessValue[0].Contains("\""))
                                            {
                                                var arr = attref.TextString.Replace(thicknessValue[0], MyTools.InchThk2mm(thicknessValue[0].Trim()).ToString()).Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                                                this.THK = double.Parse(arr[0]);
                                                this.Grade = arr[1];
                                            }
                                            else
                                            {
                                                this.THK = double.Parse(thicknessValue[0]);
                                                this.Grade = thicknessValue[1];
                                            }
                                        }

                                        break;
                                    case "PIECENAME":
                                        this.PartName = attref.TextString;
                                        break;
                                    case "QTY":
                                        this.Quantity = int.Parse(attref.TextString.Split(new char[] { '=' }, StringSplitOptions.RemoveEmptyEntries)[1]);
                                        break;
                                    case "ASS":
                                        this.PartType = attref.TextString;
                                        break;
                                    case "AREA":
                                        this.Area = double.Parse(attref.TextString);
                                        break;
                                }
                            }
                        }
                        #endregion
                        Region re = cur as Region;

                        Dictionary<int, double[]> dicAngle = new Dictionary<int, double[]>();
                        for (int i = 0; i <= 360; i++)
                        {
                            blk.MyRotate(cur.GeometricExtents.MinPoint, (Math.PI / 180) * i);
                            cur.MyRotate(cur.GeometricExtents.MinPoint, (Math.PI / 180) * i);
                            dicAngle[i] = new double[] { cur.Area / cur.GetSize()[0] * cur.GetSize()[1], cur.GetSize()[0], cur.GetSize()[1] };
                        }
                        this.OptimizationAngle = dicAngle.First(c => c.Value[0] == dicAngle.Values.Max(p => p[0])).Key;
                        this.centerPnt = re.GetCpnt();
                    }
                }
            }
        }
    }
    public class RawMaterial
    {
        public double THK { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public string MaterialGrade { get; set; }
        public string PartsType { get; set; }
        public Point3d LeftBtmPnt { get; set; }
        public Point3d RightTopPnt { get; set; }
        public ObjectId Oid { get; set; }
        public List<NCParts> NestedParts { get; set; }
        public Database acDb { get; set; }
        public double Scale { get; set; }
        public double partGaps { get; set; }
        public double platePartsGAP { get; set; }

        public RawMaterial(NCParts firstparts, Database db, Point3d _pnt, double _l, double _w, double _Sc)
        {
            this.acDb = db; this.LeftBtmPnt = _pnt; this.Scale = _Sc; this.Length = _l/_Sc; this.Width = _w / _Sc;

            this.partGaps = 20 / this.Scale; this.platePartsGAP = 20 / this.Scale;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                using (Transaction trans1 = db.TransactionManager.StartTransaction())
                {
                    if (this.Scale == 1)
                    {
                        this.Oid = MTO.InsertBlockref(this.acDb.CurrentSpaceId, firstparts.BlkName, "Material_Plate", this.LeftBtmPnt, new Scale3d(1, 1, 1), 0);
                    }
                    else this.Oid = MTO.InsertBlockref(this.acDb.CurrentSpaceId, firstparts.BlkName, "Material_Plate_Inch", this.LeftBtmPnt, new Scale3d(1, 1, 1), 0);
                    trans1.Commit();
                }

                BlockReference curRawMatlBlockref = trans.GetObject(this.Oid, OpenMode.ForWrite) as BlockReference;
                Point3d partInsPnt = new Point3d(this.LeftBtmPnt.X + (200 / (this.Scale * 10)), this.LeftBtmPnt.Y + (200 / (this.Scale * 10)), this.LeftBtmPnt.Z);//第一个零件插入点；
                if (curRawMatlBlockref != null)
                {
                    this.MaterialGrade = firstparts.Grade;
                    this.THK = firstparts.THK;
                    this.PartsType = firstparts.PartType;
                    foreach (ObjectId item in curRawMatlBlockref.AttributeCollection)
                    {
                        AttributeReference attref = trans.GetObject(item, OpenMode.ForWrite) as AttributeReference;
                        if (attref.Tag == "THK") attref.TextString = firstparts.THK.ToString();
                        if (attref.Tag == "Grade") attref.TextString = this.MaterialGrade;
                        if (attref.Tag == "Remark") attref.TextString = firstparts.PartType;
                    }
                    foreach (DynamicBlockReferenceProperty item in curRawMatlBlockref.DynamicBlockReferencePropertyCollection)
                    {
                        if (item.PropertyName == "Length") item.Value = this.Length ;
                        if (item.PropertyName == "Width") item.Value = this.Width ;
                    }
                    curRawMatlBlockref.ResetBlock();

                    ObjectId textStyleid;

                    DBText txt = new DBText()
                    {
                        Layer = firstparts.BlkName,
                        Height = 1700 / this.Scale,
                        Position = new Point3d(this.LeftBtmPnt.X - (1000 / this.Scale), this.LeftBtmPnt.Y - (3000 / this.Scale), this.LeftBtmPnt.Z)
                    };
                    if (this.Scale == 25.4)
                    {
                        txt.TextString = $"{MyTools.mmThk2inch(firstparts.THK.ToString())} ({firstparts.Grade})";
                    }
                    else txt.TextString = $"{firstparts.THK.ToString()} ({firstparts.Grade})";

                    //设置字体
                    TextStyleTable st = this.acDb.TextStyleTableId.GetObject(OpenMode.ForWrite) as TextStyleTable;
                    if (st.Has("Arial")) textStyleid = st["Arial"]; else textStyleid = st["Standard"];
                    txt.TextStyleId = textStyleid;
                    //设置字体
                    this.acDb.AddToModelSpace(txt);
                }
                trans.Commit();
            }
        }

        public void NestedNCPartsToRaw(List<NCParts> allparts)
        {
            //20190502
            Point3d partIns = new Point3d(this.LeftBtmPnt.X + this.platePartsGAP, this.LeftBtmPnt.Y + this.platePartsGAP, this.LeftBtmPnt.Z);
            double lengthLeft = this.Length; double widthLeft = this.Width;
            bool fisrtPart = true; NCParts nextNCPart = null; NCParts curNCPart = null;
            this.NestedParts = new List<NCParts>();
            NextPartLocation nextPartLoc = NextPartLocation.PutRight;//决定当前的零件如何摆放
            do
            {
                if (nextNCPart == null) curNCPart = allparts[allparts.Count - 1]; else curNCPart = nextNCPart;
                using (Transaction trans = this.acDb.TransactionManager.StartTransaction())
                {
                    Entity ent_partBo = trans.GetObject(curNCPart.ShapeObjectID, OpenMode.ForWrite) as Entity;
                    Entity ent_partLabel = trans.GetObject(curNCPart.LabelObjectID, OpenMode.ForWrite) as Entity;
                    Point3d curpartLeftBtm = Point3d.Origin;
                    #region //roation parr and move lable
                    if (curNCPart.OptimizationAngle != 0 && curNCPart.OptimizationAngle != 0.5 * Math.PI)//倾斜的零件
                    {
                        ent_partBo.MyRotate(curNCPart.centerPnt, Math.PI * 2 - curNCPart.OptimizationAngle);
                        double l = ent_partBo.GeometricExtents.MaxPoint.X - ent_partBo.GeometricExtents.MinPoint.X;
                        double w = ent_partBo.GeometricExtents.MaxPoint.Y - ent_partBo.GeometricExtents.MinPoint.Y;
                        Point3d newcen = new Point3d(
                            0.5 * (ent_partBo.GeometricExtents.MaxPoint.X + ent_partBo.GeometricExtents.MinPoint.X),
                            0.5 * (ent_partBo.GeometricExtents.MaxPoint.Y + ent_partBo.GeometricExtents.MinPoint.Y),
                            0.5 * (ent_partBo.GeometricExtents.MaxPoint.Z + ent_partBo.GeometricExtents.MinPoint.Z)
                            );
                        curpartLeftBtm = new Point3d(newcen.X - 0.5 * l, newcen.Y - 0.5 * w, newcen.Z);
                        ent_partLabel.MyRotate(curNCPart.centerPnt, Math.PI * 2 - curNCPart.OptimizationAngle);
                    }
                    else
                    {
                        if (curNCPart.OptimizationAngle == Math.PI * 0.5)//数值的零件
                        {
                            ent_partBo.MyRotate(curNCPart.centerPnt, Math.PI * 1.5);
                            curpartLeftBtm = new Point3d(
                                curNCPart.centerPnt.X - 0.5 * curNCPart.Length,
                                curNCPart.centerPnt.Y - 0.5 * curNCPart.Width, curNCPart.centerPnt.Z);
                            ent_partLabel.MyRotate(curNCPart.centerPnt, Math.PI * 1.5);
                        }
                        else
                        {
                            curpartLeftBtm = new Point3d(
                                curNCPart.centerPnt.X - 0.5 * curNCPart.Length,
                                curNCPart.centerPnt.Y - 0.5 * curNCPart.Width, curNCPart.centerPnt.Z);
                        }
                    }
                    #endregion
                    ent_partBo.MyMove(partIns, curpartLeftBtm);
                    ent_partLabel.MyMove(partIns, curpartLeftBtm);
                    trans.Commit();
                    #region //clac raw plate left size
                    if (nextPartLoc == NextPartLocation.PutUp)
                    {
                        lengthLeft = this.Length - partGaps - curNCPart.Length - platePartsGAP;
                        widthLeft = widthLeft - partGaps - curNCPart.Width;
                    }
                    else
                    {
                        if (fisrtPart)
                        {
                            widthLeft = this.Width-(platePartsGAP + curNCPart.Width + partGaps);
                            lengthLeft = this.Length-(partGaps + curNCPart.Length + platePartsGAP);
                            fisrtPart = false;
                        }
                        else lengthLeft = lengthLeft-(partGaps + curNCPart.Length);
                    }
                    #endregion

                    if (allparts.Count > 1)
                    {
                        #region //if can nested find next parts and its information
                        var fitLengthParts = allparts.Where(c => c.Length <= lengthLeft && c.Width <= curNCPart.Width && c.ShapeObjectID != curNCPart.ShapeObjectID).ToList();
                        var fitWidthParts = allparts.Where(c => c.Width <= widthLeft && c.ShapeObjectID != curNCPart.ShapeObjectID).ToList();
                        if (fitLengthParts.Count == 0 && fitWidthParts.Count == 0)
                        {
                            nextNCPart = null;
                        }
                        else
                        {
                            if (fitLengthParts.Count > 0 && fitWidthParts.Count > 0)
                            {
                                nextNCPart = fitLengthParts.OrderByDescending(c => c.Width).ToList()[0];
                                partIns = new Point3d(partIns.X + partGaps + curNCPart.Length, partIns.Y, partIns.Z);
                                nextPartLoc = NextPartLocation.PutRight;
                            }
                            else
                            {
                                if (fitWidthParts.Count > 0)
                                {
                                    nextNCPart = fitWidthParts.OrderByDescending(c => c.Length).ToList()[0];
                                    partIns = new Point3d(this.LeftBtmPnt.X + platePartsGAP,
                                        this.LeftBtmPnt.Y + platePartsGAP + partGaps + (this.Width - widthLeft)
                                        , this.LeftBtmPnt.Z);
                                    nextPartLoc = NextPartLocation.PutUp;
                                }
                                else
                                {
                                    nextNCPart = fitLengthParts.OrderByDescending(c => c.Width).ToList()[0];
                                    partIns = new Point3d(partIns.X + partGaps + curNCPart.Length,
                                        partIns.Y, partIns.Z);
                                    nextPartLoc = NextPartLocation.PutRight;
                                }
                            }
                        }
                        #endregion
                    }
                    else nextNCPart = null;
                }
                allparts.Remove(curNCPart);
                this.NestedParts.Add(curNCPart);
            } while (nextNCPart != null);
        }
    }
    public enum NextPartLocation
    {
        PutRight, PutUp
    }
}