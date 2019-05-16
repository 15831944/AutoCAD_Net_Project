using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KFWH_CMD
{
    public class myAutoCADcmdAttribute : Attribute
    {
        public string  HelpString { get; set; }
        public string IcoFileName { get; set; }
        /// <summary>
        /// 特性
        /// </summary>
        /// <param name="_str1">帮助文件内容，提示内容</param>
        /// <param name="_str2">图标的文件名称，不含路径</param>
        public myAutoCADcmdAttribute(string _str1,string _str2)
        {
            this.HelpString = _str1;
            this.IcoFileName = _str2;
        }
    }
}
