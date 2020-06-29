using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using MultiExcelMultiDoor.Excel;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MultiExcelMultiDoor
{

    public class Command 

    {
        private List<List<Point2d>> hatchList = new List<List<Point2d>>();
        private List<ObjectIdCollection> objco = new List<ObjectIdCollection>();
        [CommandMethod("HZ" )]
        public void GenerateDoor()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            var doc = Application.DocumentManager.MdiActiveDocument;
            //Application.DocumentManager.MdiActiveDocument.CommandEnded += Command_Ended;           
            try
            {
                var path = ExcelOption.ChooseSingleExcel();
                if (string.IsNullOrEmpty(path))
                {
                    return;
                }
                //第一步：找出边缘点放到数组里 如果一个不是四周都与其他点接壤，这个点就是边界点。先找出来放到数组里。
               // var pBorders = ExcelOption.ReadExcel(path);
               var arr = ExcelOption.ReadCsv(path);
               var pBorders = ExcelOption.GetPoint2ds(arr);
                var currentListSet = new List<Point2d>();
                this.hatchList.Clear();
                //第二步：排序点
                this.SortBorderLoop(pBorders, currentListSet);
                var plinelist = new List<List<Point2d>>();

                //第三步：根据贝塞尔曲线算法根据这些点拟合曲线
                foreach (var hatch in hatchList)
                {
                    var newPoints = new List<Point2d>();
                    for (int i = 0; i < hatch.Count / 2 - 2; i++)
                    {

                        var points = DrawBezier.CalculateBezierPoints(hatch[i * 2], hatch[2 * i + 1], hatch[i * 2 + 2]);
                        newPoints.Add(hatch[i * 2]);
                        newPoints.AddRange(points);
                    }
                    newPoints.Add(hatch[hatch.Count - 2]);
                    newPoints.Add(hatch[hatch.Count - 1]);
                    newPoints.Add(hatch.Last());
                    plinelist.Add(newPoints);
                }
                this.objco.Clear();
                foreach (var plines in plinelist)
                {


                    var pline = new Polyline();
                    pline.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 0, 0);
                    foreach (var p in plines)
                    {
                        pline.AddVertexAt(pline.NumberOfVertices, p, 0, 0, 0);
                    }
                    pline.Closed = true;
                    var objectid = pline.ToSpace();
                    var col = new ObjectIdCollection();
                    col.Add(objectid);
                    this.objco.Add(col); ;

                }

                //第四步 计算填充
               doc.SendCommand("AddHatch\n");

            }
            catch (System.Exception e)
            {
                ed.WriteMessage("生成失败！" + e);
            }

        }

        
        /// <summary>
        /// 通过循环方式去排序
        /// </summary>
        private void SortBorderLoop(List<Point2d> pointsRemains, List<Point2d> pointsCurrentSet)
        {

            //初始化先增加后减少
            pointsCurrentSet.Add(pointsRemains[0]);
            pointsRemains.Remove(pointsRemains[0]);
            while (pointsRemains.Count > 0)
            {
                var temp = pointsRemains.OrderBy(p => p.GetDistanceTo(pointsCurrentSet[pointsCurrentSet.Count - 1])).FirstOrDefault();
                pointsCurrentSet.Add(temp);
                pointsRemains.Remove(temp);
            }

            var listSet = new List<Point2d> { pointsCurrentSet[0] };
            for (int i = 1; i < pointsCurrentSet.Count; i++)
            {
                if (pointsCurrentSet[i].GetDistanceTo(pointsCurrentSet[i - 1]) < 1.5)
                {
                    listSet.Add(pointsCurrentSet[i]);
                }
                else
                {
                    this.hatchList.Add(listSet);
                    listSet = new List<Point2d> {pointsCurrentSet[i]};
                }
            }
            this.hatchList.Add(listSet);
        }
        /// <summary>
        /// 通过递归去排序
        /// </summary>
        /// <param name="pointsRemains"></param>
        /// <param name="PointsCurrentSet"></param>
        private void SortPBordersold(List<Point2d> pointsRemains,List<Point2d> PointsCurrentSet)
        {
            try
            {
                if (pointsRemains.Count == 0) return;

                //初次放点先放第一个点
                if (PointsCurrentSet.Count == 0)
                {
                    PointsCurrentSet.Add(pointsRemains[0]);
                    pointsRemains.RemoveAt(0);
                }               
                var ptemp = pointsRemains.Where(p => p.GetDistanceTo(PointsCurrentSet[PointsCurrentSet.Count - 1]) < 1.5).ToList();
                if (ptemp.Any())
                {                   
                    var pToAdd = ptemp.MinBy(p => p.GetDistanceTo(PointsCurrentSet[PointsCurrentSet.Count - 1]));
                    PointsCurrentSet.Add(pToAdd);
                    pointsRemains.Remove(pToAdd);
                    SortPBordersold(pointsRemains, PointsCurrentSet);
                }
                else
                {                   
                    hatchList.Add(PointsCurrentSet);                  
                    var list = new List<Point2d>();
                    SortPBordersold(pointsRemains, list);
                }
            }
            catch (System.Exception e)
            {
                throw e;
            }
            


        }
        
        
        
        [CommandMethod("AddHatch")]
        public void AddHatch()
        {
            Database db = HostApplicationServices.WorkingDatabase;
                    
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                Hatch hatch = new Hatch();//创建填充对象
                hatch.ToSpace();
                hatch.PatternScale = 0.5;//设置填充图案的比例                                                          
                string patterName = DBHelper.CurrentPattern;
                //根据上面的填充图案名创建图案填充，类型为预定义,与边界关联
                hatch.CreateHatch(HatchPatternType.PreDefined, patterName, true);
                for (int i = 0; i < this.objco.Count; i++)
                {                   
                    if (i==0)
                    {
                        //为填充添加外边界（正六边形）
                        hatch.AppendLoop(HatchLoopTypes.Outermost, objco[i]);
                    }
                    else
                    {
                        //为填充添加外边界（正六边形）
                        hatch.AppendLoop(HatchLoopTypes.Default, objco[i]);
                    }                                                                            
                }
                hatch.EvaluateHatch(true);//计算并显示填充对象
                trans.Commit();//提交更改

            }
        }
    }
}
