using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using bbOffice.Common;
using word = Aspose.Words;

namespace bbOffice.Aspose.Words
{
    public partial class FormAsposeWords : FormBase, IFiles
    {
        public FormAsposeWords()
        {
            InitializeComponent();
            this.init();
            this.initEvent();
        }

        private void init()
        {
            //doc
            this.ucFilesAndButtons1.TxbTemplateFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Words.doc";
            this.ucFilesAndButtons1.TxbExportFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Words.export.doc";

            //docx
            this.ucFilesAndButtons1.TxbTemplateFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Words.docx";
            this.ucFilesAndButtons1.TxbExportFileName.Text = Application.StartupPath + @"\Files\TestForAspose.Words.export.docx";

            this.ucFilesAndButtons1.TxbImageFilePath.Text = Application.StartupPath + @"\Files\huluwa.png";
            this.ucFilesAndButtons1.TxbPdf.Text = Application.StartupPath + @"\Files\TestForAspose.Words.export.pdf";
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
            word.Document doc = new word.Document(this.ucFilesAndButtons1.TxbTemplateFileName.Text);

            //foreach (word.Drawing.Shape shp in doc.GetChildNodes(word.NodeType.Shape, true))
            //{
            //    // iterate through each Paragraph node of Shape
            //    foreach (word.Paragraph para in shp.GetChildNodes(word.NodeType.Paragraph, true))
            //    {
            //        //Get text of Paragraph
            //        String txt = para.ToString(word.SaveFormat.Text);
            //        word.Style style = para.ParagraphFormat.Style;
            //        String StyleName = style.Name;
            //    }
            //}

            //foreach (word.Drawing.Shape shp in doc.GetChildNodes(word.NodeType.Shape, true))
            //{
            //    // iterate through each Run node of Shape
            //    foreach (word.Run run in shp.GetChildNodes(word.NodeType.Run, true))
            //    {
            //        //Get text of Run
            //        String txt = run.Text;
            //        //Get font name
            //        String fontname = run.Font.Name;

            //        //Get font size
            //        double size = run.Font.Size;
            //    }
            //}

            foreach (word.Drawing.Shape shp in doc.GetChildNodes(word.NodeType.Shape, true))
            {
                //word.Run run2 = shp.GetChildNodes(word.NodeType.Run, true).Cast<word.Run>().SingleOrDefault(pp => pp.Text == "二维码");
                //MessageBox.Show(run2.Text);
                // iterate through each Run node of Shape
                foreach (word.Run run in shp.GetChildNodes(word.NodeType.Run, true))
                {
                    //String txt = run.Text;                    
                    //String fontname = run.Font.Name;
                    //double size = run.Font.Size;
                    
                    if (run.Text == "二维码")
                    {
                        run.Text = "";
                        shp.StrokeColor = Color.Transparent;
                        word.DocumentBuilder builder = new word.DocumentBuilder(doc);
                        builder.InsertImage(this.ucFilesAndButtons1.TxbImageFilePath.Text,
                                            word.Drawing.RelativeHorizontalPosition.Margin,
                                            shp.Left,
                                            word.Drawing.RelativeVerticalPosition.Margin,
                                            shp.Top,
                                            shp.Width,
                                            shp.Height,
                                            word.Drawing.WrapType.Square);
                        break;
                    }
                }
            }

            doc.Save(this.ucFilesAndButtons1.TxbPdf.Text);
            MessageBox.Show("ok");
        }

    }
}
