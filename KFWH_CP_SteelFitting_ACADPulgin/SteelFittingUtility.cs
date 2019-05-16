using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//AutoCAD Com+
//using Autodesk.AutoCAD.Interop.Common;
//using Autodesk.AutoCAD.Interop;
//Excel Object
using Microsoft.CSharp;
//AutoCAD NetApi
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using System.Diagnostics;

//
using AutoCAD_Net_Project;
using KFWH_CMD;

namespace KFWH_CP_SteelFitting_ACADPulgin
{
    public class SteelFittingUtility
    {
        [myAutoCADcmd("Use For Steel Fitting Drawing Assign Profile Length,profile Lines must under Layer which name contains (PIPE/STIFF)", "SF_ProfilesLengthAssign.bmp")]
        [CommandMethod("SF_ProfilesLengthAssign")]//20181121 write by Nans
        public void SF_取型材长度()
        {
            Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            double sacle = 1;
            
            if (Application.GetSystemVariable("Lunits").ToString() == "4") sacle = 25.4;// 1==>Scientific;2==>Decimal;3==>Engineering;4==> Architectural;5==> Fractional
            //将用户坐标系转换成世界坐标系
            if (Application.GetSystemVariable("WORLDUCS").ToString() != "1") { ed.CurrentUserCoordinateSystem = Matrix3d.Identity; ed.Regen(); }
            //设置颜色样式
            if (Application.GetSystemVariable("CECOLOR").ToString() != "BYLAYER") Application.SetSystemVariable("CECOLOR", "BYLAYER");
            string profileSizeText = ""; double txtheight = 0; double txtAngle = 0;
            PromptSelectionOptions pso = new PromptSelectionOptions();
            PromptSelectionResult psr = null;
            //选择规格文字，必须单行文字
            pso.MessageForAdding = "Please select one Text Can Express the profile Size：";
            pso.AllowDuplicates = false; pso.SingleOnly = true;
            psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Text") }));
            if (psr.Status == PromptStatus.OK)
            {
                using (Transaction trans = acDB.TransactionManager.StartTransaction())
                {
                    BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    var obj = trans.GetObject(psr.Value[0].ObjectId, OpenMode.ForRead) as DBText;
                    profileSizeText = obj.TextString;
                    txtheight = obj.Height / 3;
                    txtAngle = obj.Rotation;
                }
            }
            // 选择型材
            pso.MessageForAdding = "Please select the Line Express Profile Routing(Line/PolyLine/Circle) Can be Selected";
            pso.AllowDuplicates = false; pso.SingleOnly = false;
            DBObjectCollection listCurves = new DBObjectCollection();
            acDB.MyCreateLayer("SteelFiting_ProfileLengthText", 160, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, false);
            psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Line,Circle,LWPOLYLINE,Arc"), new TypedValue((int)DxfCode.LayerName, "*STIFF,*PIPE") }));
            if (psr.Status == PromptStatus.OK)
            {
                try
                {
                    using (Transaction trans = acDB.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        foreach (SelectedObject item in psr.Value)
                        {
                            var obj = trans.GetObject(item.ObjectId, OpenMode.ForRead);
                            listCurves.Add(obj);
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage(ex.Message);
                }
                foreach (var item in listCurves)
                {
                    using (Transaction trans = acDB.TransactionManager.StartTransaction())
                    {
                        Point3d pnt3d = Point3d.Origin;
                        double length = 0.0d;
                        double angle = 0;
                        BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        if (item is Line)
                        {
                            length = Math.Sqrt((item as Line).Delta.X * (item as Line).Delta.X + (item as Line).Delta.Y * (item as Line).Delta.Y);
                            pnt3d = new Point3d(0.5 * ((item as Line).StartPoint.X + (item as Line).EndPoint.X), 0.5 * ((item as Line).StartPoint.Y + (item as Line).EndPoint.Y), ((item as Line).StartPoint.Z));
                            angle = (item as Line).Angle;
                        }
                        if (item is Circle)
                        {
                            length = (item as Circle).Circumference;
                            pnt3d = (item as Circle).Center;
                        }
                        if (item is Arc)
                        {
                            length = (item as Arc).Length;
                            pnt3d = new Point3d(0.5 * ((item as Arc).StartPoint.X + (item as Arc).EndPoint.X), 0.5 * ((item as Arc).StartPoint.Y + (item as Arc).EndPoint.Y), 0.5 * ((item as Arc).StartPoint.Z));
                        }
                        if (item is Polyline)
                        {
                            length = (item as Polyline).Length;
                            pnt3d = (item as Polyline).StartPoint;
                        }

                        DBText txt = new DBText()
                        {
                            TextString = acDoc.SteelFittingDwgName().Replace("-", "") + ":" + profileSizeText + "=" + (length * sacle).ToString("0.00") + "mm",
                            Rotation = angle,
                            Position = pnt3d,
                            Height = txtheight,
                            Layer = "SteelFiting_ProfileLengthText"
                        };
                        acDB.AddToModelSpace(txt);
                        trans.Commit();
                    }
                }
            }
        }

        [myAutoCADcmd("给part编号", "SF_PartsNumberAssign.bmp")]
        [CommandMethod("SF_PartsNumberAssign")]//20181122 write by Nans
        public void SF_零件编号()
        {
            Application.SetSystemVariable("OSNAPCOORD", 0);
            string partType;
            Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            string panelsNameprefix = acDoc.SteelFittingDwgName().Replace("-", "");
            //判断公制还是环境
            if (Application.GetSystemVariable("WORLDUCS").ToString() != "1") { ed.CurrentUserCoordinateSystem = Matrix3d.Identity; ed.Regen(); }
            //设置颜色样式
            if (Application.GetSystemVariable("CECOLOR").ToString() != "BYLAYER") Application.SetSystemVariable("CECOLOR", "BYLAYER");
            string sectionDetText = ""; double txtheight = 0; double rotAngle = 0;
            PromptSelectionOptions pso = new PromptSelectionOptions();
            PromptSelectionResult psr = null;
            PromptResult psrKey = null; PromptResult psrKey1 = null;
            #region//提取Section或者Detail的名称
            psrKey = ed.GetKeywords("Please Confirm you need Manually Input or MouseSelect : ", new string[] { "CursorSelect", "KeyIn" });
            if (psrKey.Status == PromptStatus.OK)
            {
                switch (psrKey.StringResult)
                {
                    case "CursorSelect":
                        pso.MessageForAdding = "Please select one Text which Contains \"DETAIL\" or \"SECTION\"";
                        pso.AllowDuplicates = false;
                        psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Text") }));
                        if (psr.Status == PromptStatus.OK)
                        {
                            using (Transaction trans = acDB.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                var obj = trans.GetObject(psr.Value[0].ObjectId, OpenMode.ForRead) as DBText;
                                // %% U DETAIL 16 - A(TYP)
                                sectionDetText = obj.TextString.ToUpper().Trim().Replace("%%U", ""); txtheight = obj.Height;
                                if (sectionDetText.Contains("DETAIL") || sectionDetText.Contains("SECTION"))
                                {
                                    if (sectionDetText.Contains("DETAIL"))
                                    {
                                        if (sectionDetText.Contains("(TYP)")) panelsNameprefix += ("/DET" + sectionDetText.Replace("DETAIL", "").Replace("(TYP)", "").Trim().Replace("-", "")); else panelsNameprefix += ("/DET" + sectionDetText.Replace("DETAIL", "").Trim().Replace("-", ""));
                                    }
                                    else
                                    {
                                        if (sectionDetText.Contains("(TYP)")) panelsNameprefix += ("/SEC" + sectionDetText.Replace("SECTION", "").Replace("(TYP)", "").Trim().Replace("-", "")); else panelsNameprefix += ("/SEC" + sectionDetText.Replace("SECTION", "").Trim().Replace("-", ""));
                                    }
                                }
                                else
                                {
                                    ed.WriteMessage("Please select one Text which Contains \"DETAIL\" or \"SECTION\"");
                                    return;
                                }
                            }
                        }
                        break;
                    case "KeyIn":
                        PromptResult ps = ed.GetString(new PromptStringOptions("Please Input a string which Identify the view "));
                        if (ps.Status == PromptStatus.OK) { sectionDetText = ps.StringResult; panelsNameprefix += ("/" + sectionDetText); } else return;
                        break;
                }
                #endregion
                #region//用户选择零件的类型
                psrKey1 = ed.GetKeywords("Please a Type as parts Type: ", new string[] { "HPlate", "TwebFace", "ProfileFlatBar", "Collar", "DBLR", "Breacket", "WebPlate", "Flange", "PlateChock", "GatePlate", "Others" });
                if (psrKey1.Status == PromptStatus.OK)
                {
                    partType = psrKey1.StringResult;
                    //ed.WriteMessage(panelsNameprefix + "-" + partType);
                    switch (partType)
                    {
                        case "HPlate":
                            panelsNameprefix += "-H";
                            break;
                        case "TwebFace":
                            panelsNameprefix += "-FC";
                            break;
                        case "ProfileFlatBar":
                            panelsNameprefix += "-F";
                            break;
                        case "Collar":
                            panelsNameprefix += "-C";
                            break;
                        case "DBLR":
                            panelsNameprefix += "-DB";
                            break;
                        case "Breacket":
                            panelsNameprefix += "-B";
                            break;
                        case "WebPlate":
                            panelsNameprefix += "-WEB";
                            break;
                        case "Flange":
                            panelsNameprefix += "-FLG";
                            break;
                        case "PlateChock":
                            panelsNameprefix += "-CH";
                            break;
                        case "GatePlate":
                            panelsNameprefix += "-GP";
                            break;
                        case "Others":
                            panelsNameprefix += "-PL";
                            break;
                    }
                }
                #endregion
                //选择代表零件的图层的形状
                pso.MessageForAdding = "Please select the Entity Express Part Shape And Material(PolyLine)Under Layer which name was \"KFWHCP_SF\" Can be Selected";
                pso.AllowDuplicates = false; pso.SingleOnly = false;
                DBObjectCollection listCurves = new DBObjectCollection();
                acDB.MyCreateLayer("SteelFiting_PartLabel", 150, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, true);
                acDB.MyCreateLayer("NC-CUTTING", 1, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, true);
                psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Circle,LWPOLYLINE"), new TypedValue((int)DxfCode.LayerName, "KFWHCP_SF") }));
                if (psr.Status == PromptStatus.OK)
                {
                    try
                    {
                        using (Transaction trans = acDB.TransactionManager.StartTransaction())
                        {
                            BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                            foreach (SelectedObject item in psr.Value)
                            {
                                var obj = trans.GetObject(item.ObjectId, OpenMode.ForRead);
                                listCurves.Add(obj);
                            }
                            trans.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(ex.Message);
                    }
                    int count = 1;
                    foreach (DBObject item in listCurves)
                    {
                        //每个零件单独处理
                        using (Transaction trans = acDB.TransactionManager.StartTransaction())
                        {
                            try
                            {
                                BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                Point3d pnt3d = new Point3d(
                                    ((item as Curve).GeometricExtents.MaxPoint.X + (item as Curve).GeometricExtents.MinPoint.X) / 2,
                                 ((item as Curve).GeometricExtents.MaxPoint.Y + (item as Curve).GeometricExtents.MinPoint.Y) / 2,
                                0
                                 );
                                string Thk = "N.A"; string qty = "0";
                                #region//获取零件材料和数量信息
                                PromptSelectionResult pselect = null;
                                if (item is Polyline)
                                {
                                    var pline = item as Polyline;
                                    pselect = ed.SelectCrossingPolygon(pline.GetAllPoints(),
                                       new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "TEXT"),
                                        new TypedValue((int)DxfCode.LayerName, "KFWHCP_SF") }));
                                    if (pselect.Status != PromptStatus.OK)
                                    {
                                        pselect = ed.SelectCrossingWindow(pline.GeometricExtents.MaxPoint, pline.GeometricExtents.MinPoint,
                                      new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "TEXT"),
                                        new TypedValue((int)DxfCode.LayerName, "KFWHCP_SF") }));
                                    }
                                }
                                else
                                {
                                    var cir = item as Circle;
                                    pselect = ed.SelectCrossingWindow(new Point3d(cir.Center.X - cir.Radius, cir.Center.Y - cir.Radius, cir.Center.Z), new Point3d(cir.Center.X + cir.Radius, cir.Center.Y + cir.Radius, cir.Center.Z),
                                       new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "TEXT"),
                                        new TypedValue((int)DxfCode.LayerName, "KFWHCP_SF") }));
                                }
                                DBObjectCollection listText = new DBObjectCollection();
                                #region //生成标签
                                if (pselect.Status == PromptStatus.OK)
                                {
                                    try
                                    {
                                        using (Transaction trans1 = acDB.TransactionManager.StartTransaction())
                                        {
                                            BlockTableRecord btr1 = trans1.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                            foreach (SelectedObject item1 in pselect.Value)
                                            {
                                                var obj = trans1.GetObject(item1.ObjectId, OpenMode.ForRead);
                                                listText.Add(obj);
                                            }
                                            foreach (var txt in listText)
                                            {
                                                if ((txt as DBText).TextString.StartsWith("M:")) Thk = (txt as DBText).TextString.Replace("M:", "").Trim(); else qty = (txt as DBText).TextString.Trim();
                                                txtheight = (txt as DBText).Height;
                                                rotAngle = (txt as DBText).Rotation;
                                            }
                                            //写入零件信息
                                            MText mtxt = new MText() { Contents = (panelsNameprefix + count) + "\\P" + Thk + "," + qty, TextHeight = 0.8 * txtheight, Location = pnt3d, Rotation = rotAngle };
                                            acDB.AddToModelSpace(mtxt);
                                            mtxt.Layer = "SteelFiting_PartLabel";
                                            mtxt.Attachment = AttachmentPoint.MiddleCenter;
                                            mtxt.BackgroundFill = true;
                                            mtxt.BackgroundScaleFactor = 1.0;
                                            mtxt.UseBackgroundColor = true;
                                            trans1.Commit();
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        ed.WriteMessage(ex.Message);
                                    }
                                    Curve partBoundary = (Curve)item.ObjectId.GetObject(OpenMode.ForWrite);
                                    partBoundary.Layer = "NC-CUTTING";
                                }
                                #endregion
                                #endregion
                                trans.Commit();
                            }
                            catch (System.Exception ex)
                            {
                                ed.WriteMessage(ex.Message);
                                trans.Abort();
                            }
                            count++;
                        }
                    }
                }
            }
        }

        [myAutoCADcmd("steel fitting的零件转化成nestix可识别的格式", "SF_CoverttoNestixParts.bmp")]
        [CommandMethod("SF_CoverttoNestixParts")]//20181207 write by Nans
        public void SF_零件转Nestix格式()
        {
            List<string> liststr = new List<string>();
            Application.SetSystemVariable("OSNAPCOORD", 0);
            Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //将用户坐标系转换成世界坐标系
            if (Application.GetSystemVariable("WORLDUCS").ToString() != "1") { ed.CurrentUserCoordinateSystem = Matrix3d.Identity; ed.Regen(); }
            //设置颜色样式
            if (Application.GetSystemVariable("CECOLOR").ToString() != "BYLAYER") Application.SetSystemVariable("CECOLOR", "BYLAYER");
            PromptSelectionOptions pso = new PromptSelectionOptions();
            acDB.MyCreateLayer("NC-CUTTING", 1, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, true);
            acDB.MyCreateLayer("NC-LABEL", 221, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, true);
            acDB.MyCreateLayer("NC-MARKING", 3, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, true);
            acDB.MyCreateLayer("NC-MARKING_TXT", 4, Autodesk.AutoCAD.Colors.ColorMethod.ByAci, true);
            pso.MessageForAdding = "Please select the Parts Under Layer which name was \"NC-CUTTING\" Can be Selected";
            pso.AllowDuplicates = false; pso.SingleOnly = false;
            var psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Circle,LWPOLYLINE"), new TypedValue((int)DxfCode.LayerName, "NC-CUTTING") }));
            DBObjectCollection listCurves = new DBObjectCollection();
            if (psr.Status == PromptStatus.OK)
            {
                try
                {
                    using (Transaction trans = acDB.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = trans.GetObject(acDB.GetModelSpaceId(), OpenMode.ForWrite) as BlockTableRecord;
                        foreach (SelectedObject item in psr.Value)
                        {
                            var obj = trans.GetObject(item.ObjectId, OpenMode.ForRead);
                            listCurves.Add(obj);
                        }
                        foreach (var item in listCurves)
                        {
                            using (Transaction trans1 = acDB.TransactionManager.StartTransaction())
                            {
                                BlockTableRecord btr1 = trans1.GetObject(acDB.GetModelSpaceId(), OpenMode.ForWrite) as BlockTableRecord;
                                //part 信息筛选
                                PromptSelectionResult pselect = null;
                                //part Marking 线型筛选
                                PromptSelectionResult pselect_MarkingLine = null;
                                //part Marking 文字筛选
                                PromptSelectionResult pselect_MarkingTxt = null;
                                #region//构件选择集
                                if (item is Polyline)
                                {
                                    var pline = item as Polyline;
                                    pselect = ed.SelectCrossingPolygon(pline.GetAllPoints(),
                                       new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "MTEXT"),
                                        new TypedValue((int)DxfCode.LayerName, "SteelFiting_PartLabel") }));
                                    pselect_MarkingLine = ed.SelectCrossingPolygon(pline.GetAllPoints(),
                                       new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Line,LWPOLYLINE,Arc") }));
                                    pselect_MarkingTxt = ed.SelectCrossingPolygon(pline.GetAllPoints(),
                                      new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "MTEXT,TEXT") }));
                                    if (pselect.Status != PromptStatus.OK)
                                    {
                                        pselect = ed.SelectCrossingWindow(pline.GeometricExtents.MaxPoint, pline.GeometricExtents.MinPoint,
                                       new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "MTEXT"),
                                        new TypedValue((int)DxfCode.LayerName, "SteelFiting_PartLabel") }));
                                    }
                                }
                                else
                                {
                                    var cir = item as Circle;
                                    pselect = ed.SelectCrossingWindow(new Point3d(cir.Center.X - cir.Radius, cir.Center.Y - cir.Radius, cir.Center.Z), new Point3d(cir.Center.X + cir.Radius, cir.Center.Y + cir.Radius, cir.Center.Z),
                                       new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "MTEXT"),
                                        new TypedValue((int)DxfCode.LayerName, "SteelFiting_PartLabel") }));
                                    pselect_MarkingLine = ed.SelectCrossingWindow(new Point3d(cir.Center.X - cir.Radius, cir.Center.Y - cir.Radius, cir.Center.Z), new Point3d(cir.Center.X + cir.Radius, cir.Center.Y + cir.Radius, cir.Center.Z),
                                      new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "Line,LWPOLYLINE,Arc") }));
                                    pselect_MarkingTxt = ed.SelectCrossingWindow(new Point3d(cir.Center.X - cir.Radius, cir.Center.Y - cir.Radius, cir.Center.Z), new Point3d(cir.Center.X + cir.Radius, cir.Center.Y + cir.Radius, cir.Center.Z),
                                      new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "MTEXT,TEXT") }));
                                }
                                #endregion
                                #region//获取零件的信息并转化成Nestix格式
                                bool boolNeedConvert = false;
                                if (pselect.Status == PromptStatus.OK)
                                {
                                    using (Transaction trans2 = acDB.TransactionManager.StartTransaction())
                                    {
                                        BlockTableRecord btr2 = trans2.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                        foreach (SelectedObject item1 in pselect.Value)
                                        {
                                            var obj = trans2.GetObject(item1.ObjectId, OpenMode.ForWrite);
                                            if (obj is MText)
                                            {
                                                if ((obj as MText).Text.Contains(":") && (obj as MText).Color.IsByLayer)
                                                {
                                                    Partinfor partinfo = new Partinfor((obj as MText).Text);
                                                    if (!liststr.Contains(partinfo.Name))
                                                    {
                                                        liststr.Add(partinfo.Name);
                                                        DBText[] txtArr = null;
                                                        if (partinfo.Mirror == "") txtArr = new DBText[5] { null, null, null, null, null }; else txtArr = new DBText[6] { null, null, null, null, null, null };
                                                        Point3d pnt = new Point3d((obj as MText).Location.X, (obj as MText).Location.Y + (obj as MText).TextHeight / 2, (obj as MText).Location.Z);
                                                        txtArr[0] = new DBText() { Height = (obj as MText).TextHeight / 4, TextString = partinfo.Name, Position = pnt, Rotation = (obj as MText).Rotation, HorizontalMode = TextHorizontalMode.TextLeft, Layer = "NC-LABEL" };

                                                        double angle = (obj as MText).Rotation;
                                                        pnt = new Point3d(pnt.X + (((obj as MText).TextHeight / 3) * Math.Sin(angle)), pnt.Y - (((obj as MText).TextHeight / 3) * Math.Cos(angle)), pnt.Z);
                                                        txtArr[1] = new DBText() { Height = (obj as MText).TextHeight / 4, TextString = partinfo.Material, Position = pnt, Rotation = (obj as MText).Rotation, HorizontalMode = TextHorizontalMode.TextLeft, Layer = "NC-LABEL" };
                                                        pnt = new Point3d(pnt.X + (((obj as MText).TextHeight / 3) * Math.Sin(angle)), pnt.Y - (((obj as MText).TextHeight / 3) * Math.Cos(angle)), pnt.Z);
                                                        txtArr[2] = new DBText() { Height = (obj as MText).TextHeight / 4, TextString = partinfo.Thickness, Position = pnt, Rotation = (obj as MText).Rotation, HorizontalMode = TextHorizontalMode.TextLeft, Layer = "NC-LABEL" };
                                                        pnt = new Point3d(pnt.X + (((obj as MText).TextHeight / 3) * Math.Sin(angle)), pnt.Y - (((obj as MText).TextHeight / 3) * Math.Cos(angle)), pnt.Z);
                                                        txtArr[3] = new DBText() { Height = (obj as MText).TextHeight / 4, TextString = partinfo.Quantity, Position = pnt, Rotation = (obj as MText).Rotation, HorizontalMode = TextHorizontalMode.TextLeft, Layer = "NC-LABEL" };
                                                        pnt = new Point3d(pnt.X + (((obj as MText).TextHeight / 3) * Math.Sin(angle)), pnt.Y - (((obj as MText).TextHeight / 3) * Math.Cos(angle)), pnt.Z);
                                                        txtArr[4] = new DBText() { Height = (obj as MText).TextHeight / 4, TextString = partinfo.Side, Position = pnt, Rotation = (obj as MText).Rotation, HorizontalMode = TextHorizontalMode.TextLeft, Layer = "NC-LABEL" };
                                                        pnt = new Point3d(pnt.X + (((obj as MText).TextHeight / 3) * Math.Sin(angle)), pnt.Y - (((obj as MText).TextHeight / 3) * Math.Cos(angle)), pnt.Z);
                                                        if (partinfo.Mirror != "")
                                                        {
                                                            txtArr[5] = new DBText() { Height = (obj as MText).TextHeight / 4, TextString = partinfo.Mirror, Position = pnt, Rotation = (obj as MText).Rotation, HorizontalMode = TextHorizontalMode.TextLeft, Layer = "NC-LABEL" };
                                                        }
                                                        acDB.AddToModelSpace(txtArr);
                                                        (obj as MText).ColorIndex = 171;
                                                        boolNeedConvert = true;
                                                    }
                                                    else
                                                    {
                                                        Line l1 = new Line(Point3d.Origin,
                                            new Point3d(((item as Curve).GeometricExtents.MaxPoint.X + (item as Curve).GeometricExtents.MinPoint.X) / 2,
                                            ((item as Curve).GeometricExtents.MaxPoint.Y + (item as Curve).GeometricExtents.MinPoint.Y) / 2,
                                            ((item as Curve).GeometricExtents.MaxPoint.Z + (item as Curve).GeometricExtents.MinPoint.Z) / 2));
                                                        l1.Layer = "0";
                                                        l1.ColorIndex = 5;
                                                        acDB.AddToModelSpace(l1);
                                                    }
                                                }
                                            }
                                        }
                                        trans2.Commit();
                                    }
                                }
                                else
                                {
                                    Line l1 = new Line(Point3d.Origin,
                                        new Point3d(((item as Curve).GeometricExtents.MaxPoint.X + (item as Curve).GeometricExtents.MinPoint.X) / 2,
                                        ((item as Curve).GeometricExtents.MaxPoint.Y + (item as Curve).GeometricExtents.MinPoint.Y) / 2,
                                        ((item as Curve).GeometricExtents.MaxPoint.Z + (item as Curve).GeometricExtents.MinPoint.Z) / 2));
                                    l1.Layer = "0";
                                    l1.ColorIndex = 5;
                                    acDB.AddToModelSpace(l1);
                                }
                                #endregion
                                //写打开
                                (item as Curve).UpgradeOpen();
                                if (!(item as Curve).Color.IsByLayer) (item as Curve).ColorIndex = 256;

                                #region//转换marking线
                                if (pselect_MarkingLine.Status == PromptStatus.OK && boolNeedConvert)
                                {
                                    using (Transaction trans2 = acDB.TransactionManager.StartTransaction())
                                    {
                                        BlockTableRecord btr2 = trans2.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                        foreach (SelectedObject item1 in pselect_MarkingLine.Value)
                                        {
                                            var obj = trans2.GetObject(item1.ObjectId, OpenMode.ForWrite) as Curve;
                                            if (obj.Layer != "NC-CUTTING") obj.Layer = "NC-MARKING";
                                            if (obj.Color.EntityColor.ColorMethod != ColorMethod.ByLayer) obj.ColorIndex = 3;
                                        }
                                        trans2.Commit();
                                    }
                                }
                                #endregion
                                #region//文字转换
                                if (pselect_MarkingTxt.Status == PromptStatus.OK && boolNeedConvert)
                                {
                                    using (Transaction trans2 = acDB.TransactionManager.StartTransaction())
                                    {
                                        BlockTableRecord btr2 = trans2.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                                        foreach (SelectedObject item1 in pselect_MarkingTxt.Value)
                                        {
                                            var obj = trans2.GetObject(item1.ObjectId, OpenMode.ForWrite) as Entity;
                                            if ((obj as Entity).Layer != "KFWHCP_SF" && obj.Layer != "TEXT" && obj.Layer != "SteelFiting_PartLabel" && obj.Layer != "NC-LABEL")
                                            {
                                                obj.Layer = "NC-MARKING_TXT";
                                                if (obj.Color.EntityColor.ColorMethod != ColorMethod.ByLayer) obj.ColorIndex = 4;
                                            }
                                        }
                                        trans2.Commit();
                                    }
                                }
                                #endregion
                                trans1.Commit();
                            }
                        }
                        trans.Commit();
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }

        [myAutoCADcmd("重新编号", "SF_AutoNumbering.bmp")]
        [CommandMethod("SF_AutoNumbering")]//20181211 write by Nans
        public void SF_零件重新编号()
        {
            Application.SetSystemVariable("OSNAPCOORD", 0);
            Database acDB = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Database;
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            //将用户坐标系转换成世界坐标系
            if (Application.GetSystemVariable("WORLDUCS").ToString() != "1") { ed.CurrentUserCoordinateSystem = Matrix3d.Identity; ed.Regen(); }
            //设置颜色样式
            if (Application.GetSystemVariable("CECOLOR").ToString() != "BYLAYER") Application.SetSystemVariable("CECOLOR", "BYLAYER");
            PromptSelectionOptions pso = new PromptSelectionOptions();

            pso.MessageForAdding = "Please select the all PartsName LabelUnder Layer which name was \"SteelFiting_PartLabel\" (MTEXT)Can be Selected";
            pso.AllowDuplicates = false; pso.SingleOnly = false;
            var psr = ed.GetSelection(pso, new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "MTEXT"), new TypedValue((int)DxfCode.LayerName, "SteelFiting_PartLabel") }));
            DBObjectCollection listPartLabel = new DBObjectCollection();
            if (psr.Status == PromptStatus.OK)
            {
                try
                {
                    using (Transaction trans = acDB.TransactionManager.StartTransaction())
                    {
                        BlockTableRecord btr = trans.GetObject(acDB.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        foreach (SelectedObject item in psr.Value)
                        {
                            var obj = trans.GetObject(item.ObjectId, OpenMode.ForRead);
                            listPartLabel.Add(obj);
                        }
                        int intH = 1; int intFC = 1; int intF = 1; int intC = 1; int intDB = 1; int intWEB = 1; int intCH = 1; int intFLG = 1; int intB = 1; int intGP = 1; int intpl = 1;
                        List<string> listH = new List<string>();
                        List<string> listFC = new List<string>();
                        List<string> listF = new List<string>();
                        List<string> listC = new List<string>();
                        List<string> listDB = new List<string>();
                        List<string> listWEB = new List<string>();
                        List<string> listCH = new List<string>();
                        List<string> listFLG = new List<string>();
                        List<string> listB = new List<string>();
                        List<string> listGP = new List<string>();
                        List<string> listpl = new List<string>();
                        for (int i = 0; i < listPartLabel.Count; i++)
                        {
                            var item = trans.GetObject(listPartLabel[i].ObjectId, OpenMode.ForWrite) as MText;
                            //  "-H";"-FC";"-F"/"-C";/-DB/-WEB/-CH/-FLG/-B
                            var arr = item.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                            if (arr.Length > 0)
                            {
                                if (arr[0].Contains("-H")) intH = NewMethod("-H", intH, listH, item, arr);
                                if (arr[0].Contains("-FC")) intFC = NewMethod("-FC", intFC, listFC, item, arr);
                                if (arr[0].Contains("-F") && !arr[0].Contains("-FC") && !arr[0].Contains("-FLG")) intF = NewMethod("-F", intF, listF, item, arr);
                                if (arr[0].Contains("-DB")) intDB = NewMethod("-DB", intDB, listDB, item, arr);
                                if (arr[0].Contains("-C") && !arr[0].Contains("-CH")) intC = NewMethod("-C", intC, listC, item, arr);
                                if (arr[0].Contains("-WEB")) intWEB = NewMethod("-WEB", intWEB, listWEB, item, arr);
                                if (arr[0].Contains("-CH")) intCH = NewMethod("-CH", intCH, listCH, item, arr);
                                if (arr[0].Contains("-FLG")) intFLG = NewMethod("-FLG", intFLG, listFLG, item, arr);
                                if (arr[0].Contains("-B")) intB = NewMethod("-B", intB, listB, item, arr);
                                if (arr[0].Contains("-GP")) intGP = NewMethod("-GP", intGP, listGP, item, arr);
                                if (arr[0].Contains("-PL")) intpl = NewMethod("-PL", intpl, listpl, item, arr);
                            }
                        }
                        trans.Commit();
                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ed.WriteMessage(ex.Message);
                }
            }
        }

        [myAutoCADcmd("导出型材信息，图纸不可以打开，必须保持关闭状态", "SF_ExportOutProfileReport.bmp")]
        [CommandMethod("SF_ExportOutProfileReport")]//20181220
        public void SF_导出型材信息()
        {
            FrmExportOutData frm = new FrmExportOutData();
            Application.ShowModelessDialog(frm);
            //Application.ShowModalDialog(frm);
        }

        private static int NewMethod(string str, int intH, List<string> listH, MText item, string[] arr)
        {
            if (listH.Contains(arr[0].Substring(0, arr[0].IndexOf(str))))
            {
                item.Contents = arr[0].Substring(0, arr[0].IndexOf(str)) + str + intH + "\\P" + arr[1];
                intH++;
            }
            else
            {
                intH = 1;
                item.Contents = arr[0].Substring(0, arr[0].IndexOf(str)) + str + intH + "\\P" + arr[1];
                listH.Add(arr[0].Substring(0, arr[0].IndexOf(str)));
                intH++;
            }
            return intH;
        }
        [myAutoCADcmd("导出零件信息，图纸不可以打开，必须保持关闭状态", "SF_ExportNestixPartList.bmp")]
        [CommandMethod("SF_ExportNestixPartList")]
        public void SF_导出零件清单()
        {
            FrmNestixPartsListExportOutcs frm = new FrmNestixPartsListExportOutcs();
            //Application.ShowModelessDialog(frm);
            Application.ShowModalDialog(frm);
        }
        public struct Partinfor
        {
            public string Name { get; set; }
            public string Material { get; set; }
            public string Thickness { get; set; }
            public string Quantity { get; set; }
            public string Side { get; set; }
            public string Mirror { get; set; }
            public Partinfor(string MtextContents)
            {
                //H33001 / SEC33C - H1\P4.5(A),Q: 1
                //PON/D28E(P)-SEC12-TW2-F1\P1"(ABS/DH36),P:1
                var arr1 = MtextContents.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                // "(", ")", "," 
                var arr = arr1[1].Split(new string[] { "(", ")", "," }, StringSplitOptions.RemoveEmptyEntries);

                this.Name = "N: " + arr1[0];
                this.Material = "M: " + arr[1];
                this.Thickness = "T: " + MyUtility.InchThk2mm(arr[0].Trim());
                if (arr.Length == 4)
                {
                    this.Quantity = "Q: " + arr[2].Split(':')[1];
                    this.Side = "SI: " + arr[2].Split(':')[0];
                    this.Mirror = "MI: " + arr[3].Split(':')[1];
                }
                else
                {
                    this.Quantity = "Q: " + arr[2].Split(':')[1];
                    this.Side = "SI: " + arr[2].Split(':')[0];
                    this.Mirror = "";
                }
            }
        }
    }
}
