using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using bbOffice.Common;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;  

namespace bbOffice.NPOI
{
    public partial class FormNPOI : FormBase, IFiles
    {
        public FormNPOI()
        {
            InitializeComponent();
            this.init();
            this.initEvent();
        }

        private void init()
        {
            this.ucFilesAndButtons1.TxbTemplateFileName.Text = Application.StartupPath + @"\Files\TestForNPOI.xls";
            this.ucFilesAndButtons1.TxbExportFileName.Text = Application.StartupPath + @"\Files\TestForNPOI.export.xls";
            this.ucFilesAndButtons1.TxbImageFilePath.Text = Application.StartupPath + @"\Files\huluwa.png";
        }

        private void initEvent()
        {
            this.ucFilesAndButtons1.BtnProcess.Click += (s, e) => this.process();
            this.ucFilesAndButtons1.BtnOpenTemplateFile.Click += (s, e) =>
            {
                if (System.IO.File.Exists(this.ucFilesAndButtons1.TxbTemplateFileName.Text))
                {
                    System.Diagnostics.Process.Start(this.ucFilesAndButtons1.TxbTemplateFileName.Text);
                }
                else
                {
                    MessageBox.Show("文件不存在.");
                }
            };
            this.ucFilesAndButtons1.BtnOpenExportFile.Click += (s, e) =>
            {
                if (System.IO.File.Exists(this.ucFilesAndButtons1.TxbExportFileName.Text))
                {
                    System.Diagnostics.Process.Start(this.ucFilesAndButtons1.TxbExportFileName.Text);
                }
                else
                {
                    MessageBox.Show("文件不存在.");
                }
            };
        }


        #region 接口实现
        OfficesFiles IFiles.SetFiles()
        {
            return new OfficesFiles()
            {
                TemplateFileName = this.ucFilesAndButtons1.TxbTemplateFileName.Text,
                ExportFileName = this.ucFilesAndButtons1.TxbExportFileName.Text,
                ImageFilePath = this.ucFilesAndButtons1.TxbImageFilePath.Text
            };
        }
        #endregion

        private void process()
        {
            using (FileStream fs = File.Open(this.ucFilesAndButtons1.TxbTemplateFileName.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                HSSFWorkbook wk = new HSSFWorkbook(fs);
                ISheet sheet = wk.GetSheetAt(0);
                sheet.GetRow(0).GetCell(0).SetCellValue("123");

                ////可行
                ////HSSFPatriarch pat = sheet.getDrawingPatriarch();
                ////List children = pat.getChildren();
                //HSSFPatriarch pat = (HSSFPatriarch)sheet.DrawingPatriarch;
                ////create the anchor
                //HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0, 0, 0, 4, 5);
                ////load the picture and get the picture index in the workbook
                //HSSFPicture picture = (HSSFPicture)pat.CreatePicture(anchor, LoadImage(this.ucFilesAndButtons1.TxbImageFilePath.Text, wk));

                ////不可用
                //XSSFDrawing drawing = sheet.CreateDrawingPatriarch() as XSSFDrawing;
                //IClientAnchor anchor1 = drawing.CreateAnchor(0, 0, 0, 0, 0, 0, 4, 5);
                //XSSFTextBox textbox1 = drawing.CreateTextbox(anchor1) as XSSFTextBox;
                //textbox1.SetText("123456798");

                ISheet sheet2 = wk.CreateSheet(System.DateTime.Now.ToString("yyyyMMdd HHmmss"));
                IRow row = sheet2.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                using (FileStream filess = new FileStream(this.ucFilesAndButtons1.TxbExportFileName.Text, FileMode.Create, FileAccess.Write))
                {
                    wk.Write(filess);
                }
            }
            MessageBox.Show("ok");
        }

        public int LoadImage(string path, HSSFWorkbook wb)
        {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[file.Length];
            file.Read(buffer, 0, (int)file.Length);
            return wb.AddPicture(buffer, PictureType.JPEG);
        }
        
    }
}
