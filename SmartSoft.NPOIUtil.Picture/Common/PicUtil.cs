using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmartSoft.NPOIUtil.Picture.Exceptions;
using SmartSoft.NPOIUtil.Picture.Interface;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SmartSoft.NPOIUtil.Picture.Common
{
    public class PicUtil : IPicUtil
    {
        #region initialize
        public const string prefix = "菜鸡你咋又出错了，还是怎么低级的错误：";
        #endregion

        #region Drawing
        /// <summary>
        /// Drawing
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetname"></param>
        /// <param name="picPath"></param>
        /// <param name="format"></param>
        /// <param name="dx1"></param>
        /// <param name="dy1"></param>
        /// <param name="dx2"></param>
        /// <param name="dy2"></param>
        /// <param name="col1"></param>
        /// <param name="row1"></param>
        /// <param name="col2"></param>
        /// <param name="row2"></param>
        public void Drawing(string filePath,string sheetname, string picPath, PictureType format, int ulx, int uly, int brx, int bry,bool isResize = false,double scaleX = 0, double scaleY = 0, int dx1 = 0, int dy1 = 0, int dx2 = 0, int dy2 = 0)
        {
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Open);

                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);

                XSSFSheet sheet = (XSSFSheet)xssfworkbook.GetSheet(sheetname ?? xssfworkbook.GetSheetName(0));


                byte[] bytes = File.ReadAllBytes(picPath);

                int pictureIdx = xssfworkbook.AddPicture(bytes, format);

                XSSFDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();

                XSSFClientAnchor anchor = new XSSFClientAnchor(dx1, dy1, dx2, dy2, ulx, uly, brx, bry);

                XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);
                if (isResize)
                {
                    if (scaleX == 0 && scaleY == 0)
                        pict.Resize();
                    else if (scaleX != 0 && scaleY == 0)
                        pict.Resize(scaleX);
                    else
                        pict.Resize(scaleX, scaleY);
                }

                FileStream file = new FileStream(filePath, FileMode.Create);

                xssfworkbook.Write(file);

                xssfworkbook.Close();
                fs.Close();
                fs.Dispose();
                file.Close();
                file.Dispose();
            }
            catch (Exception ex)
            {

                throw new PictureException ($"{prefix}{ex.Message}");
            }
        }
        #endregion
    }
}
