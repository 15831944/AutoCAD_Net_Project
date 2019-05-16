using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Text.RegularExpressions;
//AutoCAD Com+
//Excel Object
using Microsoft.CSharp;


namespace AutoCAD_Net_Project
{

    public partial class MTO
    {
        private readonly dynamic acadApp = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication;
        [Autodesk.AutoCAD.Runtime.CommandMethod("MTO_ProfileLength")]
        public void MTO骨材取长度()
        {
            //AcadLayer ly = acadApp.ActiveDocument.ActiveLayer;
            //this.CheckLayer("KFWH-MTO");
            acadApp.ActiveDocument.SetVariable("OSNAPCOORD", 0);
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem = Autodesk.AutoCAD.Geometry.Matrix3d.Identity;
            if (acadApp.ActiveDocument.GetVariable("WORLDUCS") != 1)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem = Autodesk.AutoCAD.Geometry.Matrix3d.Identity;
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Regen();
            }
            double scale = 1;
            if (acadApp.ActiveDocument.GetVariable("Lunits") == 4) scale = 25.4;// 1==>Scientific;2==>Decimal;3==>Engineering;4==> Architectural;5==> Fractional
            string profileSize = "";
            string length = "";
            double height = 0.0;
            object insPnt = null;
            double totalLength = 0;
            #region  //get size ifnormation
            for (int i = acadApp.ActiveDocument.SelectionSets.Count - 1; i >= 0; i--)
            {
                if (acadApp.ActiveDocument.SelectionSets.Item(i).Name == "MySS_MyGetLengthCmd")
                {
                    acadApp.ActiveDocument.SelectionSets.Item(i).Delete();
                }
            }
            dynamic acadSS = acadApp.ActiveDocument.SelectionSets.Add("MySS_MyGetLengthCmd");
            try
            {
                acadSS = acadApp.ActiveDocument.SelectionSets.Add(DateTime.Now.ToString());
                Int16[] ft = { -4, 0, 0, 0, -4 };
                object[] fd = { "<or", "Dimension", "Text", "MText", "or>" };
                acadApp.ActiveDocument.Utility.Prompt("Please select a text or dimension which contains profile size:");
                acadSS.SelectOnScreen(ft, fd);
                foreach (dynamic ent in acadSS)
                {
                    switch ((string)ent.ObjectName)
                    {
                        case "AcDbRotatedDimension":
                            var dimension = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            profileSize = dimension.TextOverride;
                            height = dimension.TextHeight;
                            break;
                        case "AcDbText":
                            var txt = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            profileSize = txt.TextString;
                            height = txt.Height;
                            break;
                        case "AcDbMText":
                            var mtxt = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            profileSize = mtxt.TextString;
                            height = mtxt.Height;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                acadApp.ActiveDocument.Utility.Prompt(ex.Message);
            }
            finally { acadSS.Clear(); }
            #endregion
            if (height == 0) height = 100.0d;
            #region//GetLength
            try
            {
                Int16[] ft = new Int16[2] { 0, 8 };
                int i = 0;
                object[] fd = { "Line,LWPOLYLINE, Circle,PolyLine", "Blk*" };
                acadApp.ActiveDocument.Utility.Prompt("Please select profile Line:");
                acadSS.SelectOnScreen(ft, fd);
                List<object> listExistingText = new List<object>();
                foreach (dynamic ent in acadSS)
                {
                    dynamic lengText=null;
                    switch ((string)ent.ObjectName)
                    {
                        case "AcDbLine":
                            var l1 = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            length = (Math.Round(Math.Sqrt(l1.Delta[0] * l1.Delta[0] + l1.Delta[1] * l1.Delta[1]), 4)*scale).ToString();
                            insPnt = (object)(acadApp.ActiveDocument.Utility.PolarPoint(l1.StartPoint, l1.Angle, l1.Length / 2));
                            lengText = acadApp.ActiveDocument.ModelSpace.AddText($"{profileSize}={length}", insPnt, height);
                            lengText.Rotation = l1.Angle;
                            lengText.Layer = l1.Layer;
                            break;
                        case "AcDbCircle":
                            var cir = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            length = (Math.Round(cir.Circumference, 4)*scale).ToString();
                            insPnt = cir.Center;
                            lengText = acadApp.ActiveDocument.ModelSpace.AddText($"{profileSize}={length}", insPnt, height);
                            lengText.Layer = cir.Layer;
                            break;
                        case "AcDbPolyline":
                            var pline = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            length = (Math.Round(pline.Length, 4) * scale).ToString();
                            insPnt = pline.Coordinate[1];
                            if (insPnt.GetType().IsArray)
                            {
                                var Pnts = (double[])insPnt;
                                lengText = acadApp.ActiveDocument.ModelSpace.AddText($"{profileSize}={length}", new double[] { Pnts[0], Pnts[1], 0 }, height);
                                lengText.Layer = pline.Layer;
                            }
                            break;
                        case "AcDb3dPolyline":
                            var pline_3D = acadApp.ActiveDocument.ObjectIdToObject(ent.ObjectID);
                            length = (scale*Math.Round(pline_3D.Length, 4)).ToString();
                            insPnt = pline_3D.Coordinate[1];
                            if (insPnt.GetType().IsArray)
                            {
                                var Pnts = (double[])insPnt;
                                lengText = acadApp.ActiveDocument.ModelSpace.AddText($"{profileSize}={string.Format("{0:000.00}", length)}", new double[] { Pnts[0], Pnts[1], 0 }, height);
                                lengText.Layer = pline_3D.Layer;
                        
                            }
                            break;
                    }
                    if (lengText != null)
                    {
                        if (this.CheckTextAlreadyInsert(listExistingText, lengText))
                        {
                            lengText.Delete();
                            i++;
                            acadApp.ActiveDocument.ModelSpace.AddLine(acadApp.ActiveDocument.ModelSpace.Origin, insPnt);
                        }
                        else
                        {
                            listExistingText.Add(lengText);
                            lengText.ScaleFactor = 0.7;
                            lengText.StyleName = "Standard";
                        }
                        totalLength += double.Parse(length);
                    }
                }
                if (i > 0)
                {
                    acadApp.ActiveDocument.Utility.Prompt($"Total Length for Select lines was {totalLength}mm,Total {i} reduplicate Profile Line, please check !");
                }
                else acadApp.ActiveDocument.Utility.Prompt($"Total Length for Select lines was {totalLength}mm.");
            }
            catch (Exception ex)
            {
                acadApp.ActiveDocument.Utility.Prompt(ex.Message);
            }
            finally { acadSS.Delete(); }
            #endregion
        }
        private void CheckLayer(string layerName)
        {
            int i = 0;
            var curLayer = acadApp.ActiveDocument.ActiveLayer;
            foreach (dynamic item in acadApp.ActiveDocument.Layers)
            {
                if (item.Name == layerName)
                {
                    i++;
                }
            }
            dynamic ly = null;
            if (i == 0)
            {
                ly = acadApp.ActiveDocument.Layers.Add(layerName);
            }
            else
            {
                ly = acadApp.ActiveDocument.Layers.Item(layerName);
            }
            acadApp.ActiveDocument.ActiveLayer = ly;
        }

        [Autodesk.AutoCAD.Runtime.CommandMethod("MTO_AutoNumbering")]
        public void MTO零件重新编号()
        {
            acadApp.ActiveDocument.Utility.Prompt("Please select all mtoSize Label");
            for (int i = acadApp.ActiveDocument.SelectionSets.Count - 1; i >= 0; i--)
            {
                if (acadApp.ActiveDocument.SelectionSets.Item(i).Name == "MySSAutoNumbering")
                {
                    acadApp.ActiveDocument.SelectionSets.Item(i).Delete();
                }
            }
            dynamic ss = acadApp.ActiveDocument.SelectionSets.Add("MySSAutoNumbering");
            try
            {
                Int16[] ft = new Int16[1];//注意数据类型的不一致，
                object[] fd = new object[1];
                ft[0] = 0;
                fd[0] = "INSERT";
                ss.SelectOnScreen(ft, fd);
                if (ss != null)
                {
                    int i_plate = 1; int i_hold = 1; int i_BKT = 1; int i_Web = 1; int i = 0;
                    foreach (dynamic ent in ss)
                    {
                        if (ent.EffectiveName == "PartsInfor")
                        {
                            foreach (var item in ent.GetAttributes())
                            {
                                if (item.TagString == "PIECENAME")
                                {
                                    string temp = item.TextString;
                                    if (temp.Contains("-PL")) { item.TextString = temp.Substring(0, temp.IndexOf("-PL")) + "-PL" + i_hold; i_hold++; }
                                    if (temp.Contains("-TW")) { item.TextString = temp.Substring(0, temp.IndexOf("-TW")) + "-TW" + i_Web; i_Web++; }
                                    if (temp.Contains("-BK")) { item.TextString = temp.Substring(0, temp.IndexOf("-BK")) + "-BK" + i_BKT; i_BKT++; }
                                    if (temp.Contains("-H")) { item.TextString = temp.Substring(0, temp.IndexOf("-H")) + "-H" + i_plate; i_plate++; }
                                }
                            }
                        }
                        i++;
                    }
                    acadApp.ActiveDocument.Utility.Prompt($"total {i - 1} parts update Naming index ! \n");
                }
            }
            catch (Exception ex)
            {

                acadApp.ActiveDocument.Utility.Prompt(ex.Message);
            }
            finally
            {
                ss.Delete();
            }
        }
        public bool CheckTextAlreadyInsert(List<dynamic> listText, dynamic text)
        {
            var txt = listText.Where(c => c.InsertionPoint[0] == text.InsertionPoint[0] && c.InsertionPoint[1] == text.InsertionPoint[1] && c.InsertionPoint[2] == text.InsertionPoint[2] && c.Layer == text.Layer && c.Rotation == text.Rotation && c.TextString == text.TextString).ToList();
            if (txt.Count > 0)
            {
                return true;
            }
            else return false;
        }
    }
}
