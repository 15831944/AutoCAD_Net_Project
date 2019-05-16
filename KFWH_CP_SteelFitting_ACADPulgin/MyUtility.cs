using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//AutoCAD NetApi
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using System.Text.RegularExpressions;
using System.IO;

namespace KFWH_CP_SteelFitting_ACADPulgin
{
    public static class MyUtility
    {
        public static void MyCreateLayer(this Database acdb, string layerName, short colorIndex, ColorMethod corlorMethod, bool needPrint)
        {
            using (Transaction trans = acdb.TransactionManager.StartTransaction())
            {
                LayerTable layerTable = trans.GetObject(acdb.LayerTableId, OpenMode.ForRead) as LayerTable;
                if (!layerTable.Has(layerName))
                {
                    LayerTableRecord ltr = new LayerTableRecord() { Name = layerName, Color = Color.FromColorIndex(corlorMethod, colorIndex) };
                    ltr.IsPlottable = needPrint;
                    layerTable.UpgradeOpen();
                    layerTable.Add(ltr);
                    trans.AddNewlyCreatedDBObject(ltr, true);
                }
                else
                {
                    layerTable.UpgradeOpen();
                    LayerTableRecord ltr = trans.GetObject(layerTable[layerName], OpenMode.ForWrite) as LayerTableRecord;
                    ltr.Color = Color.FromColorIndex(corlorMethod, colorIndex);
                    ltr.IsPlottable = needPrint;
                }
                trans.Commit();
            }
        }
        public static string SteelFittingDwgName(this Document acDoc)
        {
            System.Text.RegularExpressions.Regex regDwgName = new System.Text.RegularExpressions.Regex(@"-H\d{3}-\d{0,2}");
            if (regDwgName.IsMatch(acDoc.Name))
            {
                var mas = regDwgName.Matches(acDoc.Name);
                return mas[0].Value;
            }
            else return "N.A.";
        }
        public static double InchThk2mm(string inchValue)
        {
            double thk = 0.0d;
            inchValue = inchValue.Trim();
            if (inchValue.EndsWith("\""))
            {
                switch (inchValue.Replace("\"", ""))
                {
                    case "3/16": thk = 4.76; break;
                    case "1/4": thk = 6.35; break;
                    case "5/16": thk = 8.0; break;
                    case "3/8": thk = 9.5; break;
                    case "7/16": thk = 11.11; break;
                    case "1/2": thk = 12.7; break;
                    case "9/16": thk = 14.29; break;
                    case "5/8": thk = 16.0; break;
                    case "11/16": thk = 17.5; break;
                    case "3/4": thk = 19.0; break;
                    case "13/16": thk = 20.64; break;
                    case "7/8": thk = 22.23; break;
                    case "1": thk = 25.4; break;
                    case "1 1/16": thk = 27.0; break;
                    case "1 1/8": thk = 28.6; break;
                    case "1 1/4": thk = 31.75; break;
                    case "1 3/8": thk = 35.0; break;
                    case "1 1/2": thk = 38.0; break;
                    case "1 5/8": thk = 41.28; break;
                    case "1 3/4": thk = 44.50; break;
                    case "1 7/8": thk = 47.63; break;
                    case "2": thk = 50.8; break;
                    case "2 1/8": thk = 54.0; break;
                    case "2 1/4": thk = 57.15; break;
                    case "2 3/8": thk = 60.33; break;
                    case "2 1/2": thk = 63.5; break;
                    case "2 3/4": thk = 69.85; break;
                    case "3": thk = 76.2; break;
                }
                return thk;
            }
            else return double.Parse(inchValue);
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
    }
    public static partial class Tools
    {
        /// <summary>
        /// 判断字符串是否为数字
        /// </summary>
        /// <param name="value">字符串</param>
        /// <returns>如果字符串为数字，返回true，否则返回false</returns>
        public static bool IsNumeric(this string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }

        /// <summary>
        /// 判断字符串是否为整数
        /// </summary>
        /// <param name="value">字符串</param>
        /// <returns>如果字符串为整数，返回true，否则返回false</returns>
        public static bool IsInt(this string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }

        /// <summary>
        /// 获取当前.NET程序所在的目录
        /// </summary>
        /// <returns>返回当前.NET程序所在的目录</returns>
        public static string GetCurrentPath()
        {
            var moudle = System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0];
            return System.IO.Path.GetDirectoryName(moudle.FullyQualifiedName);
        }

        /// <summary>
        /// 判断字符串是否为空或空白
        /// </summary>
        /// <param name="value">字符串</param>
        /// <returns>如果字符串为空或空白，返回true，否则返回false</returns>
        public static bool IsNullOrWhiteSpace(this string value)
        {
            if (value == null) return false;
            return string.IsNullOrEmpty(value.Trim());
        }

        /// <summary>
        /// 获取模型空间的ObjectId
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回模型空间的ObjectId</returns>
        public static ObjectId GetModelSpaceId(this Database db)
        {
            return SymbolUtilityServices.GetBlockModelSpaceId(db);
        }

        /// <summary>
        /// 获取图纸空间的ObjectId
        /// </summary>
        /// <param name="db"></param>
        /// <returns>返回图纸空间的ObjectId</returns>
        public static ObjectId GetPaperSpaceId(this Database db)
        {
            return SymbolUtilityServices.GetBlockPaperSpaceId(db);
        }

        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId</returns>
        public static ObjectId AddToModelSpace(this Database db, Entity ent)
        {
            ObjectId entId;//用于返回添加到模型空间中的实体ObjectId
                           //定义一个指向当前数据库的事务处理，以添加直线
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //以读方式打开块表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                //以写方式打开模型空间块表记录.
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                entId = btr.AppendEntity(ent);//将图形对象的信息添加到块表记录中
                trans.AddNewlyCreatedDBObject(ent, true);//把对象添加到事务处理中
                trans.Commit();//提交事务处理
            }
            return entId; //返回实体的ObjectId
        }

        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToModelSpace(this Database db, params Entity[] ents)
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

        /// <summary>
        /// 将实体添加到图纸空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到图纸空间中的实体ObjectId</returns>
        public static ObjectId AddToPaperSpace(this Database db, Entity ent)
        {
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockPaperSpaceId(db), OpenMode.ForWrite);
            ObjectId id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
        }

        /// <summary>
        /// 将实体添加到图纸空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到图纸空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToPaperSpace(this Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockPaperSpaceId(db), OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }

        /// <summary>
        /// 将实体添加到当前空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到当前空间中的实体ObjectId</returns>
        public static ObjectId AddToCurrentSpace(this Database db, Entity ent)
        {
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            ObjectId id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
        }

        /// <summary>
        /// 将实体添加到当前空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到当前空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToCurrentSpace(this Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }

        /// <summary>
        /// 将字符串形式的句柄转化为ObjectId
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="handleString">句柄字符串</param>
        /// <returns>返回实体的ObjectId</returns>
        public static ObjectId HandleToObjectId(this Database db, string handleString)
        {
            Handle handle = new Handle(Convert.ToInt64(handleString, 16));
            ObjectId id = db.GetObjectId(false, handle, 0);
            return id;
        }

        /// <summary>
        /// 亮显实体
        /// </summary>
        /// <param name="ids">要亮显的实体的Id集合</param>
        public static void HighlightEntities(this ObjectIdCollection ids)
        {
            if (ids.Count == 0) return;
            var trans = ids[0].Database.TransactionManager;
            foreach (ObjectId id in ids)
            {
                Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    ent.Highlight();
                }
            }
        }

        /// <summary>
        /// 亮显选择集中的实体
        /// </summary>
        /// <param name="selectionSet">选择集</param>
        public static void HighlightEntities(this SelectionSet selectionSet)
        {
            if (selectionSet == null) return;
            ObjectIdCollection ids = new ObjectIdCollection(selectionSet.GetObjectIds());
            ids.HighlightEntities();
        }

        /// <summary>
        /// 取消亮显实体
        /// </summary>
        /// <param name="ids">实体的Id集合</param>
        public static void UnHighlightEntities(this ObjectIdCollection ids)
        {
            if (ids.Count == 0) return;
            var trans = ids[0].Database.TransactionManager;
            foreach (ObjectId id in ids)
            {
                Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    ent.Unhighlight();
                }
            }
        }

        /// <summary>
        /// 将字符串格式的点转换为Point3d格式
        /// </summary>
        /// <param name="stringPoint">字符串格式的点</param>
        /// <returns>返回对应的Point3d</returns>
        public static Point3d StringToPoint3d(this string stringPoint)
        {
            string[] strPoint = stringPoint.Trim().Split(new char[] { '(', ',', ')' }, StringSplitOptions.RemoveEmptyEntries);
            double x = Convert.ToDouble(strPoint[0]);
            double y = Convert.ToDouble(strPoint[1]);
            double z = Convert.ToDouble(strPoint[2]);
            return new Point3d(x, y, z);
        }

        /// <summary>
        /// 获取数据库对应的文档对象
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回数据库对应的文档对象</returns>
        public static Document GetDocument(this Database db)
        {
            return Application.DocumentManager.GetDocument(db);
        }

        /// <summary>
        /// 根据数据库获取命令行对象
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回命令行对象</returns>
        public static Editor GetEditor(this Database db)
        {
            return Application.DocumentManager.GetDocument(db).Editor;
        }

        /// <summary>
        /// 在命令行输出信息
        /// </summary>
        /// <param name="ed">命令行对象</param>
        /// <param name="message">要输出的信息</param>
        public static void WriteMessage(this Editor ed, object message)
        {
            ed.WriteMessage(message.ToString());
        }

        /// <summary>
        /// 在命令行输出信息，信息显示在新行上
        /// </summary>
        /// <param name="ed">命令行对象</param>
        /// <param name="message">要输出的信息</param>
        public static void WriteMessageWithReturn(this Editor ed, object message)
        {
            ed.WriteMessage("\n" + message.ToString());
        }
    }
    public static class EntTools
    {
        /// <summary>
        /// 移动实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="sourcePt">移动的源点</param>
        /// <param name="targetPt">移动的目标点</param>
        public static void Move(this ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于移动实体的矩阵
            Vector3d vector = targetPt.GetVectorTo(sourcePt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            //以写的方式打开id表示的实体对象
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);//对实体实施移动
            ent.DowngradeOpen();//为防止错误，切换实体为读的状态
        }

        /// <summary>
        /// 移动实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="sourcePt">移动的源点</param>
        /// <param name="targetPt">移动的目标点</param>
        public static void Move(this Entity ent, Point3d sourcePt, Point3d targetPt)
        {
            if (ent.IsNewObject) // 如果是还未被添加到数据库中的新实体
            {
                // 构建用于移动实体的矩阵
                Vector3d vector = targetPt.GetVectorTo(sourcePt);
                Matrix3d mt = Matrix3d.Displacement(vector);
                ent.TransformBy(mt);//对实体实施移动
            }
            else // 如果是已经添加到数据库中的实体
            {
                ent.ObjectId.Move(sourcePt, targetPt);
            }
        }

        /// <summary>
        /// 复制实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="sourcePt">复制的源点</param>
        /// <param name="targetPt">复制的目标点</param>
        /// <returns>返回复制实体的ObjectId</returns>
        public static ObjectId Copy(this ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于复制实体的矩阵
            Vector3d vector = targetPt.GetVectorTo(sourcePt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            //获取id表示的实体对象
            Entity ent = (Entity)id.GetObject(OpenMode.ForRead);
            //获取实体的拷贝
            Entity entCopy = ent.GetTransformedCopy(mt);
            //将复制的实体对象添加到模型空间
            ObjectId copyId = id.Database.AddToModelSpace(entCopy);
            return copyId; //返回复制实体的ObjectId
        }

        /// <summary>
        /// 复制实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="sourcePt">复制的源点</param>
        /// <param name="targetPt">复制的目标点</param>
        /// <returns>返回复制实体的ObjectId</returns>
        public static ObjectId Copy(this Entity ent, Point3d sourcePt, Point3d targetPt)
        {
            ObjectId copyId;
            if (ent.IsNewObject) // 如果是还未被添加到数据库中的新实体
            {
                //构建用于复制实体的矩阵
                Vector3d vector = targetPt.GetVectorTo(sourcePt);
                Matrix3d mt = Matrix3d.Displacement(vector);
                //获取实体的拷贝
                Entity entCopy = ent.GetTransformedCopy(mt);
                //将复制的实体对象添加到模型空间
                copyId = ent.Database.AddToModelSpace(entCopy);
            }
            else
            {
                copyId = ent.ObjectId.Copy(sourcePt, targetPt);
            }
            return copyId; //返回复制实体的ObjectId
        }

        /// <summary>
        /// 旋转实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="basePt">旋转基点</param>
        /// <param name="angle">旋转角度</param>
        public static void Rotate(this ObjectId id, Point3d basePt, double angle)
        {
            Matrix3d mt = Matrix3d.Rotation(angle, Vector3d.ZAxis, basePt);
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);
            ent.DowngradeOpen();
        }

        /// <summary>
        /// 旋转实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="basePt">旋转基点</param>
        /// <param name="angle">旋转角度</param>
        public static void Rotate(this Entity ent, Point3d basePt, double angle)
        {
            if (ent.IsNewObject)
            {
                Matrix3d mt = Matrix3d.Rotation(angle, Vector3d.ZAxis, basePt);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.Rotate(basePt, angle);
            }
        }

        /// <summary>
        /// 缩放实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="basePt">缩放基点</param>
        /// <param name="scaleFactor">缩放比例</param>
        public static void Scale(this ObjectId id, Point3d basePt, double scaleFactor)
        {
            Matrix3d mt = Matrix3d.Scaling(scaleFactor, basePt);
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);
            ent.DowngradeOpen();
        }

        /// <summary>
        /// 缩放实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="basePt">缩放基点</param>
        /// <param name="scaleFactor">缩放比例</param>
        public static void Scale(this Entity ent, Point3d basePt, double scaleFactor)
        {
            if (ent.IsNewObject)
            {
                Matrix3d mt = Matrix3d.Scaling(scaleFactor, basePt);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.Scale(basePt, scaleFactor);
            }
        }

        /// <summary>
        /// 镜像实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="mirrorPt1">镜像轴的第一点</param>
        /// <param name="mirrorPt2">镜像轴的第二点</param>
        /// <param name="eraseSourceObject">是否删除源对象</param>
        /// <returns>返回镜像实体的ObjectId</returns>
        public static ObjectId Mirror(this ObjectId id, Point3d mirrorPt1, Point3d mirrorPt2, bool eraseSourceObject)
        {
            Line3d miLine = new Line3d(mirrorPt1, mirrorPt2);//镜像线
            Matrix3d mt = Matrix3d.Mirroring(miLine);//镜像矩阵
            ObjectId mirrorId = id;
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            //如果删除源对象，则直接对源对象实行镜像变换
            if (eraseSourceObject == true)
                ent.TransformBy(mt);
            else //如果不删除源对象，则镜像复制源对象
            {
                Entity entCopy = ent.GetTransformedCopy(mt);
                mirrorId = id.Database.AddToModelSpace(entCopy);
            }
            return mirrorId;
        }

        /// <summary>
        /// 镜像实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="mirrorPt1">镜像轴的第一点</param>
        /// <param name="mirrorPt2">镜像轴的第二点</param>
        /// <param name="eraseSourceObject">是否删除源对象</param>
        /// <returns>返回镜像实体的ObjectId</returns>
        public static ObjectId Mirror(this Entity ent, Point3d mirrorPt1, Point3d mirrorPt2, bool eraseSourceObject)
        {
            Line3d miLine = new Line3d(mirrorPt1, mirrorPt2);//镜像线
            Matrix3d mt = Matrix3d.Mirroring(miLine);//镜像矩阵
            ObjectId mirrorId = ObjectId.Null;
            if (ent.IsNewObject)
            {
                //如果删除源对象，则直接对源对象实行镜像变换
                if (eraseSourceObject == true)
                    ent.TransformBy(mt);
                else //如果不删除源对象，则镜像复制源对象
                {
                    Entity entCopy = ent.GetTransformedCopy(mt);
                    mirrorId = ent.Database.AddToModelSpace(entCopy);
                }
            }
            else
            {
                mirrorId = ent.ObjectId.Mirror(mirrorPt1, mirrorPt2, eraseSourceObject);
            }
            return mirrorId;
        }

        /// <summary>
        /// 偏移实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="dis">偏移距离</param>
        /// <returns>返回偏移后的实体Id集合</returns>
        public static ObjectIdCollection Offset(this ObjectId id, double dis)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            Curve cur = id.GetObject(OpenMode.ForWrite) as Curve;
            if (cur != null)
            {
                try
                {
                    //获取偏移的对象集合
                    DBObjectCollection offsetCurves = cur.GetOffsetCurves(dis);
                    //将对象集合类型转换为实体类的数组，以方便加入实体的操作
                    Entity[] offsetEnts = new Entity[offsetCurves.Count];
                    offsetCurves.CopyTo(offsetEnts, 0);
                    //将偏移的对象加入到数据库
                    ids = id.Database.AddToModelSpace(offsetEnts);
                }
                catch
                {
                    Application.ShowAlertDialog("无法偏移！");
                }
            }
            else
                Application.ShowAlertDialog("无法偏移！");
            return ids;//返回偏移后的实体Id集合
        }

        /// <summary>
        /// 偏移实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="dis">偏移距离</param>
        /// <returns>返回偏移后的实体集合</returns>
        public static DBObjectCollection Offset(this Entity ent, double dis)
        {
            DBObjectCollection offsetCurves = new DBObjectCollection();
            Curve cur = ent as Curve;
            if (cur != null)
            {
                try
                {
                    offsetCurves = cur.GetOffsetCurves(dis);
                    Entity[] offsetEnts = new Entity[offsetCurves.Count];
                    offsetCurves.CopyTo(offsetEnts, 0);
                }
                catch
                {
                    Application.ShowAlertDialog("无法偏移！");
                }
            }
            else
                Application.ShowAlertDialog("无法偏移！");
            return offsetCurves;
        }

        /// <summary>
        /// 矩形阵列实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="numRows">矩形阵列的行数,该值必须为正数</param>
        /// <param name="numCols">矩形阵列的列数,该值必须为正数</param>
        /// <param name="disRows">行间的距离</param>
        /// <param name="disCols">列间的距离</param>
        /// <returns>返回阵列后的实体集合的ObjectId</returns>
        public static ObjectIdCollection ArrayRectang(this ObjectId id, int numRows, int numCols, double disRows, double disCols)
        {
            // 用于返回阵列后的实体集合的ObjectId
            ObjectIdCollection ids = new ObjectIdCollection();
            // 以读的方式打开实体
            Entity ent = (Entity)id.GetObject(OpenMode.ForRead);
            for (int m = 0; m < numRows; m++)
            {
                for (int n = 0; n < numCols; n++)
                {
                    // 获取平移矩阵
                    Matrix3d mt = Matrix3d.Displacement(new Vector3d(n * disCols, m * disRows, 0));
                    Entity entCopy = ent.GetTransformedCopy(mt);// 复制实体
                                                                // 将复制的实体添加到模型空间
                    ObjectId entCopyId = id.Database.AddToModelSpace(entCopy);
                    ids.Add(entCopyId);// 将复制实体的ObjectId添加到集合中
                }
            }
            ent.UpgradeOpen();// 切换实体为写的状态
            ent.Erase();// 删除实体
            return ids;// 返回阵列后的实体集合的ObjectId
        }

        /// <summary>
        /// 环形阵列实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="cenPt">环形阵列的中心点</param>
        /// <param name="numObj">在环形阵列中所要创建的对象数量</param>
        /// <param name="angle">以弧度表示的填充角度，正值表示逆时针方向旋转，负值表示顺时针方向旋转，如果角度为0则出错</param>
        /// <returns>返回阵列后的实体集合的ObjectId</returns>
        public static ObjectIdCollection ArrayPolar(this ObjectId id, Point3d cenPt, int numObj, double angle)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            Entity ent = (Entity)id.GetObject(OpenMode.ForRead);
            for (int i = 0; i < numObj - 1; i++)
            {
                Matrix3d mt = Matrix3d.Rotation(angle * (i + 1) / numObj, Vector3d.ZAxis, cenPt);
                Entity entCopy = ent.GetTransformedCopy(mt);
                ObjectId entCopyId = id.Database.AddToModelSpace(entCopy);
                ids.Add(entCopyId);
            }
            return ids;
        }

        /// <summary>
        /// 删除实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        public static void Erase(this ObjectId id)
        {
            DBObject ent = id.GetObject(OpenMode.ForWrite);
            ent.Erase();
        }

        /// <summary>
        /// 计算图形数据库模型空间中所有实体的包围框
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回模型空间中所有实体的包围框</returns>
        public static Extents3d GetAllEntsExtent(this Database db)
        {
            Extents3d totalExt = new Extents3d();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btRcd = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId entId in btRcd)
                {
                    Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;
                    totalExt.AddExtents(ent.GeometricExtents);
                }
            }
            return totalExt;
        }
    }
}

