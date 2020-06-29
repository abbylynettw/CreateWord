using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.Geometry;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices.Core;

namespace MultiExcelMultiDoor.Excel
{
    public static class ExcelOption
    {

        /// <summary>
        /// 用户选择一个表格
        /// </summary>
        /// <returns></returns>
        public static string ChooseSingleExcel()
        {
            try
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.Multiselect = false;
                fileDialog.Title = "请选择表格";
                fileDialog.Filter = "所有文件(*.csv*)|*.csv*"; //设置要选择的文件的类型

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (fileDialog.FileNames.Length > 0)
                    {
                        return fileDialog.FileNames[0];
                    }

                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return null;
            }

            return "";


        }

        public static List<List<string>> ReadCsv(string filePath) //从csv读取数据返回table
        {
            List<Point2d> point2ds = new List<Point2d>();
            List<List<string>> data = new List<List<string>>(); //用来存放读取到的值
            try
            {
                Encoding encoding = Encoding.Default; //GetType(filePath); //
                FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs, encoding);
                //记录每次读取的一行记录
                string strLine = "";
                bool isFirst = true;
                while ((strLine = sr.ReadLine()) != null)
                {
                    var tableHead = strLine.Split(',').ToList();
                    data.Add(tableHead);
                }

                sr.Close();
                fs.Close();
                return data;
            }
            catch (Exception e)
            {
                return data;
            }
        }

        public static List<Point2d> GetPoint2ds(List<List<string>> data)
        {
            try
            {
                var plist = new List<Point2d>();
                //第一步：如果一个不是四周都与其他点接壤，这个点就是边界点。先找出来放到数组里。
                for (int r = 0; r < data.Count; r++)
                {
                    for (int c = 0; c < data[r].Count; c++)
                    {
                        var value = data[r][c];
                        if (value != "0" && !string.IsNullOrWhiteSpace(value))
                        {
                            if (r > 0 && c > 0 && r < data.Count && c < data[r].Count)
                            {
                                var left = data[r][c - 1];
                                var right = data[r][c + 1];
                                var up = data[r - 1][c];
                                var down = data[r + 1][c];
                                if (left == "0" || right == "0" || up == "0" || down == "0" || string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right) || string.IsNullOrWhiteSpace(up) || string.IsNullOrWhiteSpace(down))
                                {
                                    plist.Add(new Point2d(c, data.Count - r));
                                }
                            }
                        }
                    }

                }
                return plist;
            }
            catch (Exception e)
            {
                return null;
            }
         
           
        }

        /// <summary>
        /// 读取表格数据
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static List<Point2d> ReadExcel(string fileName)
        {
            List<Point2d> plist = new List<Point2d>();
            try
            {
                //创建文件流对象
                using (FileStream filesrc = File.OpenRead(fileName))
                {
                    if (string.IsNullOrEmpty(fileName))
                    {
                        return plist;
                    }

                    //工作簿对象获取Excel内容
                    IWorkbook workbook = new XSSFWorkbook(filesrc);
                    //获得工作簿里面的工作表
                    ISheet sheet = workbook.GetSheetAt(0);
                    //第一步：如果一个不是四周都与其他点接壤，这个点就是边界点。先找出来放到数组里。
                    for (int r = 0; r <= sheet.LastRowNum; r++)
                    {

                        IRow row = sheet.GetRow(r);
                        if (row != null)
                        {
                            for (int c = 0; c < row.LastCellNum; c++)
                            {
                                var value = row.GetCell(c).ToString();
                                if (value != "0" && !string.IsNullOrWhiteSpace(value))
                                {
                                    if (r > 0 && c > 0 && r < sheet.LastRowNum && c < row.LastCellNum)
                                    {
                                        var left = row.GetCell(c - 1).ToString();
                                        var right = row.GetCell(c + 1).ToString();
                                        var up = sheet.GetRow(r - 1).GetCell(c).ToString();
                                        var down = sheet.GetRow(r + 1).GetCell(c).ToString();
                                        if (left == "0" || right == "0" || up == "0" || down == "0" || string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right) || string.IsNullOrWhiteSpace(up) || string.IsNullOrWhiteSpace(down))
                                        {
                                            plist.Add(new Point2d(c, row.LastCellNum - r));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成错误！" + ex);
            }
            return plist;
        }
    }
}
