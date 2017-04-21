using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace bbOffice.Common
{
    public partial class UCFilesAndButtons : UserControl
    {
        public UCFilesAndButtons()
        {
            InitializeComponent();
        }

        public UCFilesAndButtons(IContainer container)
        {
            InitializeComponent();
            container.Add(this);
        }

        public TextBox TxbTemplateFileName
        {
            get { return txbTemplateFileName; }
        }

        public TextBox TxbExportFileName
        {
            get { return txbExportFileName; }
        }

        public TextBox TxbImageFilePath
        {
            get { return txbImageFilePath; }
        }


        public TextBox TxbPdf
        {
            get { return txbPdf; }
        }

        public Button BtnProcess
        {
            get { return btnProcess; }
        }

        public Button BtnOpenTemplateFile
        {
            get { return btnOpenTemplateFile; }
        }

        public Button BtnOpenExportFile
        {
            get { return btnOpenExportFile; }
        }


        public Button BtnOpenPdf
        {
            get { return btnOpenPdf; }
        }

        //public event BtnProcessClick     
        //{
        //    get { return btnProcess_Click; }
        //}
        

    }
}
