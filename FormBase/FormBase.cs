using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace bbOffice.Common
{
    public partial class FormBase : Form, IShowOnPage
    {

        public FormBase()
        {
            InitializeComponent();
            this.Init();
        }

        public virtual void Init()
        {
            this.InitPanel();
        }

        public virtual void InitPanel()
        {
            
        }


        #region IShowOnPage 接口实现
        Form IShowOnPage.GetForm()
        {
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.TopLevel = false;
            return this;
        }

        string IShowOnPage.GetName()
        {
            return this.Name;
        }        


        #endregion


    }
}
