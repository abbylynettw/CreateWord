using System.Collections.Generic;
using System.IO;

namespace MultiExcelMultiDoor
{
    public static class IOHelper
    {
        public static void CreateDirectory(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }
        static List<string> list = new List<string>();//定义list变量，存放获取到的路径
        /// <summary>
        /// 读取某一文件夹下的所有文件夹和文件
        /// </summary>
        /// <param name="path">文件夹路径</param>
        /// <returns></returns>
        public static List<string> GetPath(string path)
        {            
            DirectoryInfo dir = new DirectoryInfo(path);
            System.IO.FileInfo[] fil = dir.GetFiles();
            DirectoryInfo[] dii = dir.GetDirectories();
            list.Clear();
            foreach (System.IO.FileInfo f in fil)
            {                                
                list.Add(f.FullName);//添加文件的路径到列表
            }           
            return list;
        }
    }
}
