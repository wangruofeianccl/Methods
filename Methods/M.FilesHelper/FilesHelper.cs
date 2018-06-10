using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M.FilesHelper
{
    public class FilesHelper
    {
        /// <summary>
        /// 打开指定目录文件夹，返回所有文件集合
        /// </summary>
        /// <param name="path">完整文件夹路径(如："E:\icon")</param>
        /// <param name="fileType">要读取的文件夹内的文件格式（如："*.txt"）</param>
        /// <returns></returns>
        public static IList<string> openFiles(string path, string fileType = null)
        {
            try
            {
                DirectoryInfo folder = new DirectoryInfo(path);
                IList<string> _list = new List<string>();
                if (fileType != null)
                {
                    foreach (FileInfo file in folder.GetFiles(fileType))
                    {
                        // Console.WriteLine(file.FullName);
                        _list.Add(file.FullName);
                    }
                    return _list;
                }
                else
                {
                    foreach (FileInfo file in folder.GetFiles())
                    {
                        // Console.WriteLine(file.FullName);
                        _list.Add(file.FullName);
                    }
                    return _list;
                }
            }
            catch(Exception ex)
            {
                throw ex.InnerException;
            }
            
        }
    }
}
