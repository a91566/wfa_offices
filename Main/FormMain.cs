using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using bbOffice.Common;
using System.IO;
using System.Reflection;

namespace bbOffice
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            this.init();
            this.Shown +=(s, e)=> this.initForm();
        }

        private void init()
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void initForm()
        {
            var list = this.GetInterfaces();
            if (list == null || list.Count == 0)
            {
                MessageBox.Show("没有文件.");
                return;
            }
            TabControl tc = new TabControl();
            tc.Parent = this;
            tc.Dock = DockStyle.Fill;
            tc.Font = this.Font;
            tc.Appearance = TabAppearance.FlatButtons;
            foreach (IShowOnPage item in list)
            {
                TabPage tp = new TabPage();
                tc.TabPages.Add(tp);
                tc.SelectedTab = tp;
                tp.Text = item.GetName();
                Form f = item.GetForm();
                f.BackColor = this.BackColor;
                f.Parent = tp;
                f.Show();
            }
        }

        /// <summary>
        /// 加载接口
        /// </summary>
        /// <returns></returns>
        public List<IShowOnPage> GetInterfaces()
        {
            List<IShowOnPage> implementObject = new List<IShowOnPage>();
            string dir = Application.StartupPath;
            foreach (var file in Directory.GetFiles(dir, "*.dll"))
            {
                if (!file.Contains("bbOffice") || file.Contains("Common")) continue;
                //加载程序集
                var asm = Assembly.LoadFrom(file);
                //遍历程序集中的类型
                foreach (var type in asm.GetTypes())
                {
                    //如果是IzsbTest接口                    
                    if (type.GetInterfaces().Contains(typeof(IShowOnPage)))
                    {
                        //创建接口类型实例
                        IShowOnPage instance = Activator.CreateInstance(type) as IShowOnPage;
                        if (instance != null)
                        {
                            implementObject.Add(instance);
                        }
                    }
                }
            }
            return implementObject;
        }
    }
}
