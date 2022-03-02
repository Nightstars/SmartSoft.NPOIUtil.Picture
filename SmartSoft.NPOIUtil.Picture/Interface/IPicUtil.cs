using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace SmartSoft.NPOIUtil.Picture.Interface
{
    public interface IPicUtil
    {
        public void Drawing(string filePath, string sheetname, string picPath, PictureType format, int ulx, int uly, int brx, int bry, bool isResize = false, double scaleX = 0, double scaleY = 0, int dx1 = 0, int dy1 = 0, int dx2 = 0, int dy2 = 0);

        public void Drawing(string filePath, string sheetname, string picPath, PictureType format, int x, int y, double scale = 1.0, int yOffset = 0);

    }
}
