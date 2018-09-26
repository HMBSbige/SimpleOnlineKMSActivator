namespace SimpleOnlineKMSActivator
{
    partial class MainForm
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.windowsTabPage = new System.Windows.Forms.TabPage();
			this.WindowsKey = new System.Windows.Forms.TextBox();
			this.comboBox3 = new System.Windows.Forms.ComboBox();
			this.button3 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.button1 = new System.Windows.Forms.Button();
			this.label_version = new System.Windows.Forms.Label();
			this.officeTabPage = new System.Windows.Forms.TabPage();
			this.button8 = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.progressBar2 = new System.Windows.Forms.ProgressBar();
			this.button7 = new System.Windows.Forms.Button();
			this.button6 = new System.Windows.Forms.Button();
			this.button5 = new System.Windows.Forms.Button();
			this.button4 = new System.Windows.Forms.Button();
			this.textBox_OfficePath = new System.Windows.Forms.TextBox();
			this.comboBox2 = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.tabControl1.SuspendLayout();
			this.windowsTabPage.SuspendLayout();
			this.officeTabPage.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// textBox1
			// 
			this.textBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.textBox1.Location = new System.Drawing.Point(0, 235);
			this.textBox1.MaxLength = 0;
			this.textBox1.Multiline = true;
			this.textBox1.Name = "textBox1";
			this.textBox1.ReadOnly = true;
			this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.textBox1.Size = new System.Drawing.Size(324, 199);
			this.textBox1.TabIndex = 0;
			this.textBox1.WordWrap = false;
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.windowsTabPage);
			this.tabControl1.Controls.Add(this.officeTabPage);
			this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabControl1.Location = new System.Drawing.Point(0, 53);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(324, 182);
			this.tabControl1.TabIndex = 8;
			// 
			// windowsTabPage
			// 
			this.windowsTabPage.Controls.Add(this.WindowsKey);
			this.windowsTabPage.Controls.Add(this.comboBox3);
			this.windowsTabPage.Controls.Add(this.button3);
			this.windowsTabPage.Controls.Add(this.button2);
			this.windowsTabPage.Controls.Add(this.progressBar1);
			this.windowsTabPage.Controls.Add(this.linkLabel1);
			this.windowsTabPage.Controls.Add(this.button1);
			this.windowsTabPage.Controls.Add(this.label_version);
			this.windowsTabPage.Location = new System.Drawing.Point(4, 22);
			this.windowsTabPage.Name = "windowsTabPage";
			this.windowsTabPage.Padding = new System.Windows.Forms.Padding(3);
			this.windowsTabPage.Size = new System.Drawing.Size(316, 156);
			this.windowsTabPage.TabIndex = 0;
			this.windowsTabPage.Text = "Windows";
			this.windowsTabPage.UseVisualStyleBackColor = true;
			// 
			// WindowsKey
			// 
			this.WindowsKey.Location = new System.Drawing.Point(74, 63);
			this.WindowsKey.Name = "WindowsKey";
			this.WindowsKey.Size = new System.Drawing.Size(234, 21);
			this.WindowsKey.TabIndex = 13;
			// 
			// comboBox3
			// 
			this.comboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox3.FormattingEnabled = true;
			this.comboBox3.Location = new System.Drawing.Point(8, 37);
			this.comboBox3.Name = "comboBox3";
			this.comboBox3.Size = new System.Drawing.Size(300, 20);
			this.comboBox3.TabIndex = 9;
			this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectedIndexChanged);
			// 
			// button3
			// 
			this.button3.Location = new System.Drawing.Point(74, 90);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(86, 23);
			this.button3.TabIndex = 12;
			this.button3.Text = "获取激活指令";
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// button2
			// 
			this.button2.Location = new System.Drawing.Point(162, 90);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(146, 23);
			this.button2.TabIndex = 11;
			this.button2.Text = "查询 Windows 激活状态";
			this.button2.UseVisualStyleBackColor = true;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// progressBar1
			// 
			this.progressBar1.Location = new System.Drawing.Point(8, 119);
			this.progressBar1.Maximum = 3;
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.Size = new System.Drawing.Size(219, 23);
			this.progressBar1.Step = 1;
			this.progressBar1.TabIndex = 2;
			// 
			// linkLabel1
			// 
			this.linkLabel1.AutoSize = true;
			this.linkLabel1.Location = new System.Drawing.Point(3, 67);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(65, 12);
			this.linkLabel1.TabIndex = 8;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "KMS 密钥：";
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(233, 119);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 23);
			this.button1.TabIndex = 1;
			this.button1.Text = "激活";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label_version
			// 
			this.label_version.AutoSize = true;
			this.label_version.Location = new System.Drawing.Point(6, 12);
			this.label_version.Name = "label_version";
			this.label_version.Size = new System.Drawing.Size(0, 12);
			this.label_version.TabIndex = 7;
			// 
			// officeTabPage
			// 
			this.officeTabPage.Controls.Add(this.button8);
			this.officeTabPage.Controls.Add(this.label3);
			this.officeTabPage.Controls.Add(this.progressBar2);
			this.officeTabPage.Controls.Add(this.button7);
			this.officeTabPage.Controls.Add(this.button6);
			this.officeTabPage.Controls.Add(this.button5);
			this.officeTabPage.Controls.Add(this.button4);
			this.officeTabPage.Controls.Add(this.textBox_OfficePath);
			this.officeTabPage.Location = new System.Drawing.Point(4, 22);
			this.officeTabPage.Name = "officeTabPage";
			this.officeTabPage.Padding = new System.Windows.Forms.Padding(3);
			this.officeTabPage.Size = new System.Drawing.Size(316, 156);
			this.officeTabPage.TabIndex = 1;
			this.officeTabPage.Text = "Office";
			this.officeTabPage.UseVisualStyleBackColor = true;
			// 
			// button8
			// 
			this.button8.Location = new System.Drawing.Point(169, 35);
			this.button8.Name = "button8";
			this.button8.Size = new System.Drawing.Size(139, 23);
			this.button8.TabIndex = 7;
			this.button8.Text = "自动获取 Office 目录";
			this.button8.UseVisualStyleBackColor = true;
			this.button8.Click += new System.EventHandler(this.button8_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.ForeColor = System.Drawing.Color.Red;
			this.label3.Location = new System.Drawing.Point(6, 35);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(143, 12);
			this.label3.TabIndex = 6;
			this.label3.Text = "该目录下不存在 OSPP.VBS";
			// 
			// progressBar2
			// 
			this.progressBar2.Location = new System.Drawing.Point(6, 97);
			this.progressBar2.Maximum = 2;
			this.progressBar2.Name = "progressBar2";
			this.progressBar2.Size = new System.Drawing.Size(221, 23);
			this.progressBar2.Step = 1;
			this.progressBar2.TabIndex = 5;
			// 
			// button7
			// 
			this.button7.Location = new System.Drawing.Point(74, 68);
			this.button7.Name = "button7";
			this.button7.Size = new System.Drawing.Size(89, 23);
			this.button7.TabIndex = 4;
			this.button7.Text = "获取激活指令";
			this.button7.UseVisualStyleBackColor = true;
			this.button7.Click += new System.EventHandler(this.button7_Click);
			// 
			// button6
			// 
			this.button6.Location = new System.Drawing.Point(169, 68);
			this.button6.Name = "button6";
			this.button6.Size = new System.Drawing.Size(139, 23);
			this.button6.TabIndex = 3;
			this.button6.Text = "查询 Office 激活状态";
			this.button6.UseVisualStyleBackColor = true;
			this.button6.Click += new System.EventHandler(this.button6_Click);
			// 
			// button5
			// 
			this.button5.Location = new System.Drawing.Point(203, 6);
			this.button5.Name = "button5";
			this.button5.Size = new System.Drawing.Size(105, 23);
			this.button5.TabIndex = 2;
			this.button5.Text = "选择Office目录";
			this.button5.UseVisualStyleBackColor = true;
			this.button5.Click += new System.EventHandler(this.button5_Click);
			// 
			// button4
			// 
			this.button4.Location = new System.Drawing.Point(233, 97);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(75, 23);
			this.button4.TabIndex = 1;
			this.button4.Text = "激活";
			this.button4.UseVisualStyleBackColor = true;
			this.button4.Click += new System.EventHandler(this.button4_Click);
			// 
			// textBox_OfficePath
			// 
			this.textBox_OfficePath.AllowDrop = true;
			this.textBox_OfficePath.Location = new System.Drawing.Point(6, 6);
			this.textBox_OfficePath.MaxLength = 0;
			this.textBox_OfficePath.Name = "textBox_OfficePath";
			this.textBox_OfficePath.ReadOnly = true;
			this.textBox_OfficePath.Size = new System.Drawing.Size(191, 21);
			this.textBox_OfficePath.TabIndex = 0;
			this.textBox_OfficePath.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
			this.textBox_OfficePath.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBox2_DragDrop);
			this.textBox_OfficePath.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBox2_DragEnter);
			// 
			// comboBox2
			// 
			this.comboBox2.FormattingEnabled = true;
			this.comboBox2.Items.AddRange(new object[] {
            "kms.bige0.com",
            "kms.03k.org",
            "kms.luody.info",
            "kms.bige0.cn",
            "kms02.bige0.cn"});
			this.comboBox2.Location = new System.Drawing.Point(100, 20);
			this.comboBox2.Name = "comboBox2";
			this.comboBox2.Size = new System.Drawing.Size(204, 20);
			this.comboBox2.TabIndex = 4;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(17, 23);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(77, 12);
			this.label2.TabIndex = 6;
			this.label2.Text = "KMS 服务器：";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.comboBox2);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
			this.groupBox1.Location = new System.Drawing.Point(0, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(324, 53);
			this.groupBox1.TabIndex = 9;
			this.groupBox1.TabStop = false;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(324, 434);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.textBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "MainForm";
			this.Text = "SimpleOnlineKMSActivator";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.tabControl1.ResumeLayout(false);
			this.windowsTabPage.ResumeLayout(false);
			this.windowsTabPage.PerformLayout();
			this.officeTabPage.ResumeLayout(false);
			this.officeTabPage.PerformLayout();
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage windowsTabPage;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label_version;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabPage officeTabPage;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox textBox_OfficePath;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.TextBox WindowsKey;
    }
}

