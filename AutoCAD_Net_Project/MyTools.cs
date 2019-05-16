using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace AutoCAD_Net_Project
{
    public static class MyTools
    {
        public static Point3dCollection GetAllPoints(this Polyline pline)
        {
            Point3dCollection p3dc = new Point3dCollection();
            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                p3dc.Add(pline.GetPoint3dAt(i));
            }
            return p3dc;
        }
        public static void MyMove(this ObjectId entityId, Point3d originPnt, Point3d targetPnt)
        {
            Vector3d vector = targetPnt.GetVectorTo(originPnt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            var ent = entityId.GetObject(OpenMode.ForWrite) as Entity;
            ent.TransformBy(mt);
            ent.DowngradeOpen();
        }
        public static void MyMove(this Entity ent, Point3d originPnt, Point3d targetPnt)
        {
            if (ent.IsNewObject)
            {
                Vector3d vector = targetPnt.GetVectorTo(originPnt);
                Matrix3d mt = Matrix3d.Displacement(vector);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.MyMove(originPnt, targetPnt);
            }
        }
        public static void MyRotate(this ObjectId entityId, Point3d cPnt, double angle)
        {
            Matrix3d mt = Matrix3d.Rotation(angle, Vector3d.ZAxis, cPnt);
            var ent = entityId.GetObject(OpenMode.ForWrite) as Entity;
            ent.TransformBy(mt);
            ent.DowngradeOpen();
        }
        public static void MyRotate(this Entity ent, Point3d cPnt, double angle)
        {
            if (ent.IsNewObject)
            {
                Matrix3d mt = Matrix3d.Rotation(angle, Vector3d.ZAxis, cPnt);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.MyRotate(cPnt, angle);
            }
        }
        /// <summary>
        /// 导入外部文件中的块
        /// </summary>
        /// <param name="destDb">目标数据库</param>
        /// <param name="sourceFileName">包含完整路径的外部文件名</param>
        public static void ImportBlocksFromDwg(this Database destDb, string sourceFileName)
        {
            //创建一个新的数据库对象，作为源数据库，以读入外部文件中的对象
            Database sourceDb = new Database(false, true);
            try
            {
                //把DWG文件读入到一个临时的数据库中
                sourceDb.ReadDwgFile(sourceFileName, System.IO.FileShare.Read, true, null);
                //创建一个变量用来存储块的ObjectId列表
                ObjectIdCollection blockIds = new ObjectIdCollection();
                //获取源数据库的事务处理管理器
                Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;
                //在源数据库中开始事务处理
                using (Transaction myT = tm.StartTransaction())
                {
                    //打开源数据库中的块表
                    BlockTable bt = (BlockTable)tm.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);
                    //遍历每个块
                    foreach (ObjectId btrId in bt)
                    {
                        BlockTableRecord btr = (BlockTableRecord)tm.GetObject(btrId, OpenMode.ForRead, false);
                        //只加入命名块和非布局块到复制列表中
                        if (!btr.IsAnonymous && !btr.IsLayout)
                        {
                            blockIds.Add(btrId);
                        }
                        btr.Dispose();
                    }
                    bt.Dispose();
                }
                //定义一个IdMapping对象
                IdMapping mapping = new IdMapping();
                //从源数据库向目标数据库复制块表记录
                sourceDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Application.ShowAlertDialog("复制错误: " + ex.Message);
            }
            //操作完成，销毁源数据库
            sourceDb.Dispose();
        }
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

        public static double InchThk2mm(string inchValue)
        {
            inchValue = inchValue.Trim();
            double thk = 0.0d;
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
                    case "3 1/2": thk = 88.9; break;
                    case "4": thk = 101.6; break;
                }
                return thk;
            }
            else return double.Parse(inchValue);
        }
        public static string mmThk2inch(string mmValue)
        {
            string inchVal = string.Empty;
            switch (mmValue)
            {
                case "4.76": inchVal = "3/16"; break;
                case "6.35": inchVal = "1/4"; break;
                case "8": inchVal = "5/16"; break;
                case "9.5": inchVal = "3/8"; break;
                case "11.11": inchVal = "7/16"; break;
                case "12.7": inchVal = "1/2"; break;
                case "14.29": inchVal = "9/16"; break;
                case "16": inchVal = "5/8"; break;
                case "17.5": inchVal = "11/16"; break;
                case "19": inchVal = "3/4"; break;
                case "20.64": inchVal = "13/16"; break;
                case "22.23": inchVal = "7/8"; break;
                case "25.4": inchVal = "1"; break;
                case "27": inchVal = "1 1/16"; break;
                case "28.6": inchVal = "1 1/8"; break;
                case "31.75": inchVal = "1 1/4"; break;
                case "35": inchVal = "1 3/8"; break;
                case "38": inchVal = "1 1/2"; break;
                case "41.28": inchVal = "1 5/8"; break;
                case "44.5": inchVal = "1 3/4"; break;
                case "47.63": inchVal = "1 7/8"; break;
                case "50.8": inchVal = "2"; break;
                case "54": inchVal = "2 1/8"; break;
                case "57.15": inchVal = "2 1/4"; break;
                case "60.33": inchVal = "2 3/8"; break;
                case "63.5": inchVal = "2 1/2"; break;
                case "69.85": inchVal = "2 3/4"; break;
                case "76.2": inchVal = "3"; break;
                case "88.9": inchVal = "3 1/2"; break;
                case "101.6": inchVal = "4"; break;
                default: inchVal = "N.A."; break;
            }
            return inchVal + "\"";
        }
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
        public static ObjectId AddToModelSpace(this Database db, Entity ent)
        {
            ObjectId id = new ObjectId();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
            id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
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
        public static Autodesk.AutoCAD.EditorInput.Editor GetEditor(this Database db)
        {
            return Application.DocumentManager.GetDocument(db).Editor;
        }
        public static Autodesk.AutoCAD.Geometry.Point3d GetCpnt(this Entity ent)
        {
            var pntX = (ent.GeometricExtents.MaxPoint.X + ent.GeometricExtents.MinPoint.X) / 2;
            var pntY = (ent.GeometricExtents.MaxPoint.Y + ent.GeometricExtents.MinPoint.Y) / 2;
            var pntZ = (ent.GeometricExtents.MaxPoint.Z + ent.GeometricExtents.MinPoint.Z) / 2;
            return new Point3d(pntX, pntY, pntZ);
        }
        public static double[] GetSize(this Entity ent)
        {
            var L = (ent.GeometricExtents.MaxPoint.X - ent.GeometricExtents.MinPoint.X);
            var W = (ent.GeometricExtents.MaxPoint.Y - ent.GeometricExtents.MinPoint.Y);
            return new double[] { L, W };
        }

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
    }
}
