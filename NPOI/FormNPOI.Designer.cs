namespace bbOffice.NPOI
{
    partial class FormNPOI
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.ucFilesAndButtons1 = new bbOffice.Common.UCFilesAndButtons(this.components);
            this.SuspendLayout();
            // 
            // ucFilesAndButtons1
            // 
            this.ucFilesAndButtons1.BackColor = System.Drawing.Color.White;
            this.ucFilesAndButtons1.Dock = System.Windows.Forms.DockStyle.Top;
            this.ucFilesAndButtons1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ucFilesAndButtons1.Location = new System.Drawing.Point(0, 0);
            this.ucFilesAndButtons1.Margin = new System.Windows.Forms.Padding(4);
            this.ucFilesAndButtons1.Name = "ucFilesAndButtons1";
            this.ucFilesAndButtons1.Size = new System.Drawing.Size(753, 226);
            this.ucFilesAndButtons1.TabIndex = 0;
            // 
            // FormNPOI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(753, 515);
            this.Controls.Add(this.ucFilesAndButtons1);
            this.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FormNPOI";
            this.ResumeLayout(false);

        }

        #endregion

        private Common.UCFilesAndButtons ucFilesAndButtons1;

    }
}

