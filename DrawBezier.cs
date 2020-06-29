using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;

namespace MultiExcelMultiDoor
{
    /// <summary>
    /// 贝塞尔曲线算法
    /// </summary>
   static class DrawBezier
    {
        /// <summary>
        /// 计算两点之间的二次贝塞尔曲线点集。
        /// </summary>
        public static List<Point2d> CalculateBezierPoints(Point2d begin, Point2d handle, Point2d end)
        {
            // 根据两点之间的距离来确定曲线上的点数，进而确定 t 的步长（步长越小，精度越高）
            var bx = begin.X;
            var by = begin.Y;
            var ex = end.X;
            var ey = end.Y;
            var hx = handle.X;
            var hy = handle.Y;

            var bhDis = Math.Sqrt((bx - hx) * (bx - hx) + (by - hy) * (by - hy));
            var ehDis = Math.Sqrt((ex - hx) * (ex - hx) + (ey - hy) * (ey - hy));

            var distance = bhDis + ehDis;
            var count = distance / 40;

            if (count < 2)
            {
                count = 2;
            }

            // 确定步长，并计算曲线上的点集
            var step = 0.5 / count;
            var points = new List<Point2d>();

            var t = 0.0; // t∈[0,1]
            for (var i = 0; i < count; i++)
            {
                points.Add(GetBezierPointByT(begin, handle, end, t));
                t += step;
            }

            return points;
        }

        /// <summary>
        /// 根据二次贝塞尔曲线的公式，计算各个时刻的坐标值。
        /// </summary>
        private static Point2d GetBezierPointByT(Point2d begin, Point2d handle, Point2d end, double t)
        {
            var pow = Math.Pow(1 - t, 2);

            // B = (1-t)^2*P0 + 2t(1-t)P1 + t^2*P2 , t∈[0,1]
            var x = pow * begin.X + 2 * t * (1 - t) * handle.X + t * t * end.X;
            var y = pow * begin.Y + 2 * t * (1 - t) * handle.Y + t * t * end.Y;

            return new Point2d(x, y);
        }
        
    }
}
