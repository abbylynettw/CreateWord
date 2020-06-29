using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace MultiExcelMultiDoor
{
    public static class DBHelper
    {

        /// <summary>
        /// 找到最大的某个元素
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="TR"></typeparam>
        /// <param name="en"></param>
        /// <param name="evaluate"></param>
        /// <returns></returns>
        public static T MaxBy<T, TR>(this IEnumerable<T> en, Func<T, TR> evaluate) where TR : IComparable<TR>
        {
            return en.Select(t => new Tuple<T, TR>(t, evaluate(t)))
                .Aggregate((max, next) => next.Item2.CompareTo(max.Item2) > 0 ? next : max).Item1;
        }

        public static double MyGetDistanceTo(this Point2d p1,Point2d p2)
        {
            return Math.Sqrt((p2.X - p1.X) * (p2.X - p1.X) + (p2.Y - p1.Y) * (p2.Y - p1.Y));
        }
        /// <summary>
        /// 找到最小的某个元素
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="TR"></typeparam>
        /// <param name="en"></param>
        /// <param name="evaluate"></param>
        /// <returns></returns>
        public static T MinBy<T, TR>(this IEnumerable<T> en, Func<T, TR> evaluate) where TR : IComparable<TR>
        {
            return en.Select(t => new Tuple<T, TR>(t, evaluate(t)))
                .Aggregate((max, next) => next.Item2.CompareTo(max.Item2) < 0 ? next : max).Item1;
        }


        /// <summary>
        /// 将实体添加到特定空间。
        /// </summary>
        /// <param name="ent"></param>
        /// <param name="db"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public static ObjectId ToSpace(this Entity ent, Database db = null, string space = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var id = ObjectId.Null;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                id = mdlSpc.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);

                trans.Commit();
            }

            return id;
        }

        /// <summary>
        /// 调用COM的SendCommand函数
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="args">命令参数列表</param>
        public static void SendCommand(this Document doc, params string[] args)
        {
            Type AcadDocument = Type.GetTypeFromHandle(Type.GetTypeHandle(doc.GetAcadDocument()));
            try
            {
                // 通过后期绑定的方式调用SendCommand命令
                AcadDocument.InvokeMember("SendCommand", BindingFlags.InvokeMethod, null, doc.GetAcadDocument(), args);
            }
            catch // 捕获异常
            {
                return;
            }
        }

        public static string CurrentPattern
        {
            get { return Application.GetSystemVariable("HPNAME").ToString(); }
        }

        public static void CreateHatch(this Hatch hatch, HatchPatternType type, string patternName, bool associative)
        {
            Database db = HostApplicationServices.WorkingDatabase;
            hatch.SetDatabaseDefaults();
            hatch.SetHatchPattern(type, patternName);
            hatch.Associative = associative ? true : false;
        }

    }
}
