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
                const string path = @"D:\myworkspace\mydocuments\project\普瑞均胜\导出模板_出口发票.xlsx";
                const string picpath = @"D:\myworkspace\mydocuments\project\普瑞均胜\章.png";
                instance.Drawing(path, "出口发票", picpath, PictureType.PNG, 4, 28, 4.5, 12);
                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }

        }
    }
}
