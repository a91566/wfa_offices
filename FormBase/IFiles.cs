using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace bbOffice.Common
{
    public interface IFiles
    {
        OfficesFiles SetFiles();
    }

    public class OfficesFiles
    {
        /// <summary>
        /// 模板文件（绝对路径）
        /// </summary>
        public string TemplateFileName;
        /// <summary>
        /// 导出文件（绝对路径）
        /// </summary>
        public string ExportFileName;
        /// <summary>
        /// 图片路径
        /// </summary>
        public string ImageFilePath;
    }

    
}
