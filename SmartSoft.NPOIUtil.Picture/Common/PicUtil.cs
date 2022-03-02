using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SmartSoft.NPOIUtil.Picture.Exceptions;
using SmartSoft.NPOIUtil.Picture.Interface;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace SmartSoft.NPOIUtil.Picture.Common
{
    public class PicUtil : IPicUtil
    {
        #region initialize
        public const string prefix = "糟糕，出错了： ";
        #endregion

        #region Drawing
        /// <summary>
        /// Drawing
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="sheetname">Sheet名</param>
        /// <param name="picPath">图片路径</param>
        /// <param name="format">图片格式</param>
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

        #region Drawding
        /// <summary>
        /// Drawing
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetname"></param>
        /// <param name="picPath"></param>
        /// <param name="format"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="scale"></param>
        /// <exception cref="PictureException"></exception>
        public void Drawing(string filePath, string sheetname, string picPath, PictureType format, int x, int y, double scale = 1.0, int yOffset = 0)
        {
            try
            {
                FileStream fs = new FileStream(filePath, FileMode.Open);

                XSSFWorkbook xssfworkbook = new XSSFWorkbook(fs);

                XSSFSheet sheet = (XSSFSheet)xssfworkbook.GetSheet(sheetname ?? xssfworkbook.GetSheetName(0));


                byte[] bytes = File.ReadAllBytes(picPath);

                //MemoryStream ms = new MemoryStream(bytes);
                //Image Img = Bitmap.FromStream(ms, true);

                int pictureIdx = xssfworkbook.AddPicture(bytes, format);

                XSSFDrawing patriarch = (XSSFDrawing)sheet.CreateDrawingPatriarch();

                XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, x, y, x, y + yOffset);

                XSSFPicture pict = (XSSFPicture)patriarch.CreatePicture(anchor, pictureIdx);

                //IRow row = sheet.CreateRow(x);

                //ICell cell = row.CreateCell(y);

                //var cellwidth = sheet.GetColumnWidthInPixels(cell.ColumnIndex);
                //var cellheight = cell.Sheet.GetRow(cell.RowIndex).HeightInPoints / 72 * 96;

                //scale = Fixed ? scale /cellwidth : scale;

                pict.Resize(scale, 1);

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

                throw new PictureException($"{prefix}{ex.Message}");
            }
        }
        #endregion
    }
}
