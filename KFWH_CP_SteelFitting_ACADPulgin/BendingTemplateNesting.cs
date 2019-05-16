using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;
using System.IO;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Geometry;
using KFWH_CMD;

namespace SM3DShellPanelUtility
{
    public class BendingTemplateNesting
    {
        [myAutoCADcmd("外板的样板零件处理程序，注意先炸掉里面的块参照", "BTN_ConvertParts.bmp")]
        [CommandMethod("BTN_ConvertParts")]
        public void 外板样板零件处理()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            if (acDoc.Editor.CurrentUserCoordinateSystem != Matrix3d.Identity)
            {
                acDoc.Editor.CurrentUserCoordinateSystem = Matrix3d.Identity;
                acDoc.Editor.Regen();
            }
            Type comType = Type.GetTypeFromHandle(Type.GetTypeHandle(Application.AcadApplication));
            comType.InvokeMember("ZoomExtents", System.Reflection.BindingFlags.InvokeMethod, null, Application.AcadApplication, null);
            acDb.Cecolor = Color.FromColorIndex(ColorMethod.ByAci, 3);
            #region//解冻和解锁所有图层
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                LayerTable lt = trans.GetObject(acDb.LayerTableId, OpenMode.ForWrite) as LayerTable;
                foreach (ObjectId ltrid in lt)
                {
                    LayerTableRecord ltr = trans.GetObject(ltrid, OpenMode.ForWrite) as LayerTableRecord;
                    if (ltr.IsLocked) ltr.IsLocked = false;
                    if (ltr.IsFrozen) ltr.IsFrozen = false;
                    if (ltr.IsOff) ltr.IsOff = false;
                }
                trans.Commit();
            }
            #endregion
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                List<Polyline2d> listpartsExtens = new List<Polyline2d>();
                foreach (ObjectId item in btr)
                {
                    Entity ent = trans.GetObject(item, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        if (ent is Polyline2d && ent.Layer.EndsWith("DrawingExtens") && ent.ColorIndex.ToString() == "7") listpartsExtens.Add(ent as Polyline2d);
                    }
                }
                if (listpartsExtens.Count > 0)
                {
                    foreach (var item in listpartsExtens) ConvertToTargetPart(acDoc, item);
                }
                Type comThisDrawing = Type.GetTypeFromHandle(Type.GetTypeHandle(acDoc.GetAcadDocument()));
                comThisDrawing.InvokeMember("PurgeAll", System.Reflection.BindingFlags.InvokeMethod, null, acDoc.GetAcadDocument(), null);
                trans.Commit();
            }
        }
        public static bool FileInUse(string filename)
        {
            bool use = true;
            FileStream fs = null;
            try
            {
                fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.None);
                use = false;
            }
            catch (System.Exception)
            {

            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
            return use;
        }
        public static ObjectIdCollection AddToModelSpace(Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }
        public static ObjectId AddToModelSpace(Database db, Entity ent)
        {
            ObjectId id = new ObjectId();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
            id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
        }
        public void ConvertToTargetPart(Document doc, Polyline2d partExtens)//$DrawingExtens
        {
            var pnts = partExtens.GeometricExtents;
            using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
            {
                partExtens.Erase();
                trans.Commit();
            }
            List<Entity> listCurveBoundary = new List<Entity>();
            PromptSelectionResult psr = doc.Editor.SelectWindow(pnts.MaxPoint, pnts.MinPoint);
            if (psr.Status == PromptStatus.OK)
            {
                List<ObjectId> oids = new List<ObjectId>();
                foreach (SelectedObject item in psr.Value)
                {
                    using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
                    {
                        Entity ent = trans.GetObject(item.ObjectId, OpenMode.ForWrite) as Entity;
                        if (ent is Curve && (ent.ColorIndex.ToString() == "1" || ent.ColorIndex.ToString() == "3")) { ent.Layer = "0"; ent.ColorIndex = 4; }//内部marking线
                        if (ent is Curve && ent.ColorIndex.ToString() == "7" && ent.Layer != "_Auto") { ent.Layer = "0"; ent.ColorIndex = 3; listCurveBoundary.Add(ent); }//外轮廓线
                        if (ent is Curve && ent.Layer == "_Auto") { ent.Layer = "0"; ent.ColorIndex = 4; }//内部marking线
                        if (ent is DBText) { ent.Layer = "0"; ent.ColorIndex = 6; }//外轮廓线
                        if (ent is Spline)//存在样条曲线的边界线
                        {
                            var sp = ent as Spline;
                            listCurveBoundary.Remove(ent as Curve);
                            sp.Erase();
                            Curve pline = sp.ToPolylineWithPrecision(1);
                            var oid = AddToModelSpace(doc.Database, pline);
                            oids.Add(oid);
                        }
                        trans.Commit();
                    }
                }
                using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
                {
                    Entity ent = null;
                    if (oids.Count == 0)//全是直线的情况
                    {
                        Polyline pline = new Polyline();
                        Point3dCollection pnts3d = new Point3dCollection();
                        for (int i = listCurveBoundary.Count - 1; i >= 0; i--)
                        {
                            Line l = listCurveBoundary[i] as Line;
                            pnts3d.Add(l.StartPoint);
                            pnts3d.Add(l.EndPoint);
                            l.Erase();
                        }
                        pline.CreatePolyline(pnts3d);
                        pline.Layer = "0";
                        AddToModelSpace(doc.Database, pline);
                    }
                    else
                    {
                        ent = oids[0].GetObject(OpenMode.ForWrite) as Entity;
                        for (int i = 1; i < oids.Count; i++) listCurveBoundary.Add(oids[0].GetObject(OpenMode.ForWrite) as Entity);
                        Autodesk.AutoCAD.Geometry.IntegerCollection intSet = ent.JoinEntities(listCurveBoundary.ToArray());
                        for (int i = 0; i < listCurveBoundary.Count; i++) listCurveBoundary[i].Erase();
                    }
                    trans.Commit();
                }
            }
        }

        [myAutoCADcmd("外板的型材样板零件处理程序，注意先炸掉里面的块参照", "BTN_ConvertProfiles.bmp")]
        [CommandMethod("BTN_ConvertProfiles")]
        public void 外板型材样板处理()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            if (acDoc.Editor.CurrentUserCoordinateSystem != Matrix3d.Identity)
            {
                acDoc.Editor.CurrentUserCoordinateSystem = Matrix3d.Identity;
                acDoc.Editor.Regen();
            }

            Type comType = Type.GetTypeFromHandle(Type.GetTypeHandle(Application.AcadApplication));
            comType.InvokeMember("ZoomExtents", System.Reflection.BindingFlags.InvokeMethod, null, Application.AcadApplication, null);
            acDb.Cecolor = Color.FromColorIndex(ColorMethod.ByAci, 3);
            #region//解冻和解锁所有图层
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                LayerTable lt = trans.GetObject(acDb.LayerTableId, OpenMode.ForWrite) as LayerTable;
                foreach (ObjectId ltrid in lt)
                {
                    LayerTableRecord ltr = trans.GetObject(ltrid, OpenMode.ForWrite) as LayerTableRecord;
                    if (ltr.IsLocked) ltr.IsLocked = false;
                    if (ltr.IsFrozen) ltr.IsFrozen = false;
                    if (ltr.IsOff) ltr.IsOff = false;
                }
                trans.Commit();
            }
            #endregion
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                List<Polyline2d> listpartsExtensProfiles = new List<Polyline2d>();
                foreach (ObjectId item in btr)
                {
                    Entity ent = trans.GetObject(item, OpenMode.ForWrite) as Entity;
                    if (ent != null)
                    {
                        if (ent is Polyline2d && ent.Layer.EndsWith("DrawingExtens") && ent.ColorIndex.ToString() == "7") listpartsExtensProfiles.Add(ent as Polyline2d);
                    }
                }
                if (listpartsExtensProfiles.Count > 0)
                {
                    foreach (var item in listpartsExtensProfiles) ConvertToTargetProfiles(acDoc, item);
                }
                trans.Commit();

            }
            //删除各种垃圾
            using (Transaction trans = acDb.TransactionManager.StartTransaction())
            {
                BlockTable bt = trans.GetObject(acDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                foreach (ObjectId item in btr)
                {
                    Entity ent = trans.GetObject(item, OpenMode.ForWrite) as Entity;
                    if (ent != null) if (ent.Layer != "0") ent.Erase();
                }
                Type comThisDrawing = Type.GetTypeFromHandle(Type.GetTypeHandle(acDoc.GetAcadDocument()));
                comThisDrawing.InvokeMember("PurgeAll", System.Reflection.BindingFlags.InvokeMethod, null, acDoc.GetAcadDocument(), null);
                trans.Commit();
            }
        }

        private void ConvertToTargetProfiles(Document doc, Polyline2d partExtens)
        {
            var pnts = partExtens.GeometricExtents;
            //删除边界
            using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
            {
                partExtens.Erase();
                trans.Commit();
            }
            List<Entity> listCurveBoundary = new List<Entity>();
            PromptSelectionResult psr = doc.Editor.SelectWindow(pnts.MaxPoint, pnts.MinPoint);
            if (psr.Status == PromptStatus.OK)
            {
                List<ObjectId> oids = new List<ObjectId>();
                foreach (SelectedObject item in psr.Value)
                {
                    using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
                    {
                        Entity ent = trans.GetObject(item.ObjectId, OpenMode.ForWrite) as Entity;
                        if (ent.Layer == "Rulers") //删除无用的标尺
                        {
                            ent.Erase();
                        }
                        if (ent.Layer == "Markingline")//marking 线
                        {
                            ent.Layer = "0";
                            ent.ColorIndex = 4;

                        }
                        if (ent is DBText && (ent.Layer == "Elevation Plane" || ent.Layer == "MarginSymbol" || ent.Layer == "ProfileSide"))//frame文字
                        {
                            ent.Layer = "0";
                            if (ent.Layer != "ProfileSide") (ent as DBText).Height = 20;
                            ent.ColorIndex = 6;

                        }
                        if (ent is Spline)//存在样条曲线
                        {
                            var sp = ent as Spline;
                            sp.Erase();
                            Curve pline = sp.ToPolylineWithPrecision(1);
                            pline.Layer = "0";
                            pline.ColorIndex = 3;
                            var oid = AddToModelSpace(doc.Database, pline);
                            oids.Add(oid);

                        }
                        if ((ent is Line || ent is Arc) && ent.Layer == "ProfileWebOrient")//存在样条曲线
                        {
                            ent.Layer = "0";
                            ent.ColorIndex = 3;

                        }
                        trans.Commit();
                    }
                }
                if (oids.Count > 0)
                {
                    using (Transaction trans = doc.Database.TransactionManager.StartTransaction())
                    {
                        List<Polyline> allPlines = new List<Polyline>();
                        foreach (ObjectId item in oids)
                        {
                            Polyline ent = trans.GetObject(item, OpenMode.ForWrite) as Polyline;
                            allPlines.Add(ent);
                        }
                        var minY = allPlines.Min(p => p.StartPoint.Y);
                        var targetPline = allPlines.Where(p => p.StartPoint.Y == minY).ToList()[0];
                        for (int i = allPlines.Count - 1; i >= 0; i--)
                        {
                            if (allPlines[i].ObjectId != targetPline.ObjectId)
                            {
                                allPlines[i].Erase();
                                allPlines.Remove(allPlines[i]);
                            }
                        }
                        var dbs = targetPline.GetOffsetCurves(-80);
                        foreach (DBObject item in dbs)
                        {
                            (item as Entity).Layer = "0";
                            (item as Entity).ColorIndex = 3;
                            AddToModelSpace(doc.Database, item as Entity);
                        }
                        trans.Commit();
                    }
                }

            }
        }
    }
}
