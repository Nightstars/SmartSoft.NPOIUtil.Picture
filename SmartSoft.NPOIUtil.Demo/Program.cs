using NPOI.SS.UserModel;
using SmartSoft.NPOIUtil.Picture.Factory;
using System;

namespace SmartSoft.NPOIUtil.Demo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var instance = new PicUtilFactory().GetInstance();
                string path = @"D:\myworkspace\mydocuments\project\普瑞均胜\导出模板_出口发票.xlsx";
                string picpath = @"D:\myworkspace\mydocuments\project\普瑞均胜\章.png";
                instance.Drawing(path, "出口发票", picpath, PictureType.PNG, 19, 26, 25, 32);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.ReadKey();
        }
    }
}
