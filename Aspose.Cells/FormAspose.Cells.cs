using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using bbOffice;
using bbOffice.Common;
using excel = Aspose.Cells;

namespace bbOffice.Aspose.Cells
{
    public partial class FormAsposeCells : FormBase, IFiles
    {
        public FormAsposeCells()
        {
            InitializeComponent();
            this.init();
            this.initEvent();
        }

        private void init()
        {
            // xls
            this.ucFilesAndButtons1.TxbTemplateFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Cells.xls";
            this.ucFilesAndButtons1.TxbExportFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Cells.export.xls";

            // xlsx
            this.ucFilesAndButtons1.TxbTemplateFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Cells.xlsx";
            this.ucFilesAndButtons1.TxbExportFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Cells.export.xlsx";

            this.ucFilesAndButtons1.TxbImageFilePath.Text = Application.StartupPath + @"\Files\huluwa.png";
            this.ucFilesAndButtons1.TxbPdf.Text = Application.StartupPath + @"\Files\TestForAspose.Cells.export.pdf";
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
            this.ucFilesAndButtons1.BtnOpenPdf.Click += (s, e) =>
            {
                if (System.IO.File.Exists(this.ucFilesAndButtons1.TxbPdf.Text))
                {
                    //System.Diagnostics.Process.Start(this.ucFilesAndButtons1.TxbPdf.Text);
                    this.webBrowser1.Navigate(@"file:///" + this.ucFilesAndButtons1.TxbPdf.Text);
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
            excel.Workbook workbook = new excel.Workbook(this.ucFilesAndButtons1.TxbTemplateFileName.Text);
            excel.Worksheet sheet = workbook.Worksheets[0];
            int count = sheet.TextBoxes.Count();
            for (int i = 0; i < count; i++)
            {                
                excel.Drawing.TextBox t = sheet.TextBoxes[i];
                if (t.Text == "二维码")
                {
                    t.HasLine = false;
                    t.Text = "";
                    int x = sheet.Pictures.Add(5, 5, this.ucFilesAndButtons1.TxbImageFilePath.Text);
                    excel.Drawing.Picture picture = sheet.Pictures[x];
                    picture.LeftToCorner = t.LeftToCorner;
                    picture.TopToCorner = t.TopToCorner;
                    picture.Width = t.Width;
                    picture.Height = t.Height;
                    break;
                }
                
            }
            //excel.Drawing.TextBox txb = sheet.TextBoxes[0];
            //txb.HasLine = false;
            //txb.Text = "123";
            //int x = sheet.Pictures.Add(5, 5, this.ucFilesAndButtons1.TxbImageFilePath.Text);
            //excel.Drawing.Picture picture = sheet.Pictures[x];
            //picture.LeftToCorner = txb.LeftToCorner;
            //picture.TopToCorner = txb.TopToCorner;
            //picture.Width = txb.Width;
            //picture.Height = txb.Height;
            workbook.Save(this.ucFilesAndButtons1.TxbExportFileName.Text);
            workbook.Save(this.ucFilesAndButtons1.TxbPdf.Text);
            MessageBox.Show("ok");
        }
    }
}
