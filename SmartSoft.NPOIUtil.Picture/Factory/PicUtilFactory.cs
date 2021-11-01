using SmartSoft.NPOIUtil.Picture.Common;
using SmartSoft.NPOIUtil.Picture.Interface;
using System;
using System.Collections.Generic;
using System.Text;

namespace SmartSoft.NPOIUtil.Picture.Factory
{
    public class PicUtilFactory : IPicUtilFactory
    {
        public IPicUtil GetInstance()
        {
            return new PicUtil();
        }
    }
}
