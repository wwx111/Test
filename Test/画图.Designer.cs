
namespace Test
{
    partial class 画图
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


        #endregion

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.SavePicMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.隐藏显示图例ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.隐藏显示控制板ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.切换宽高比ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.按住Ctrl鼠标滚轮缩放ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dlgSavePic = new System.Windows.Forms.SaveFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.OutZoom = new System.Windows.Forms.Button();
            this.InZoom = new System.Windows.Forms.Button();
            this.RemoveDay = new System.Windows.Forms.Button();
            this.AddDay = new System.Windows.Forms.Button();
            this.MonthRight = new System.Windows.Forms.Button();
            this.MonthLeft = new System.Windows.Forms.Button();
            this.DayRight = new System.Windows.Forms.Button();
            this.DayLeft = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.tabControl1 = new DevComponents.DotNetBar.TabControl();
            this.contextMenuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.tabControl1)).BeginInit();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.SavePicMenuItem,
            this.隐藏显示图例ToolStripMenuItem,
            this.隐藏显示控制板ToolStripMenuItem,
            this.切换宽高比ToolStripMenuItem,
            this.按住Ctrl鼠标滚轮缩放ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(269, 154);
            // 
            // SavePicMenuItem
            // 
            this.SavePicMenuItem.Name = "SavePicMenuItem";
            this.SavePicMenuItem.Size = new System.Drawing.Size(268, 30);
            this.SavePicMenuItem.Text = "保存图片";
            this.SavePicMenuItem.Click += new System.EventHandler(this.SavePicMenuItem_Click);
            // 
            // 隐藏显示图例ToolStripMenuItem
            // 
            this.隐藏显示图例ToolStripMenuItem.Name = "隐藏显示图例ToolStripMenuItem";
            this.隐藏显示图例ToolStripMenuItem.Size = new System.Drawing.Size(268, 30);
            this.隐藏显示图例ToolStripMenuItem.Text = "隐藏/显示图例";
            this.隐藏显示图例ToolStripMenuItem.Click += new System.EventHandler(this.隐藏显示图例ToolStripMenuItem_Click);
            // 
            // 隐藏显示控制板ToolStripMenuItem
            // 
            this.隐藏显示控制板ToolStripMenuItem.Name = "隐藏显示控制板ToolStripMenuItem";
            this.隐藏显示控制板ToolStripMenuItem.Size = new System.Drawing.Size(268, 30);
            this.隐藏显示控制板ToolStripMenuItem.Text = "隐藏/显示控制板";
            this.隐藏显示控制板ToolStripMenuItem.Click += new System.EventHandler(this.隐藏显示控制板ToolStripMenuItem_Click);
            // 
            // 切换宽高比ToolStripMenuItem
            // 
            this.切换宽高比ToolStripMenuItem.Name = "切换宽高比ToolStripMenuItem";
            this.切换宽高比ToolStripMenuItem.Size = new System.Drawing.Size(268, 30);
            this.切换宽高比ToolStripMenuItem.Text = "切换宽高比";
            // 
            // 按住Ctrl鼠标滚轮缩放ToolStripMenuItem
            // 
            this.按住Ctrl鼠标滚轮缩放ToolStripMenuItem.Name = "按住Ctrl鼠标滚轮缩放ToolStripMenuItem";
            this.按住Ctrl鼠标滚轮缩放ToolStripMenuItem.Size = new System.Drawing.Size(268, 30);
            this.按住Ctrl鼠标滚轮缩放ToolStripMenuItem.Text = "按住Ctrl+鼠标滚轮缩放";
            // 
            // dlgSavePic
            // 
            this.dlgSavePic.Filter = "*.bmp|*.bmp|*.png|*.png|*.jpg|*.jpg|*.gif|*.gif";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.OutZoom);
            this.panel1.Controls.Add(this.InZoom);
            this.panel1.Controls.Add(this.RemoveDay);
            this.panel1.Controls.Add(this.AddDay);
            this.panel1.Controls.Add(this.MonthRight);
            this.panel1.Controls.Add(this.MonthLeft);
            this.panel1.Controls.Add(this.DayRight);
            this.panel1.Controls.Add(this.DayLeft);
            this.panel1.Location = new System.Drawing.Point(77, 157);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(107, 238);
            this.panel1.TabIndex = 1;
            this.panel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            this.panel1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseMove);
            this.panel1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseUp);
            // 
            // OutZoom
            // 
            this.OutZoom.Image = global::Test.Properties.Resources.缩小;
            this.OutZoom.Location = new System.Drawing.Point(54, 171);
            this.OutZoom.Name = "OutZoom";
            this.OutZoom.Size = new System.Drawing.Size(50, 50);
            this.OutZoom.TabIndex = 7;
            this.toolTip1.SetToolTip(this.OutZoom, "缩小图片");
            this.OutZoom.UseVisualStyleBackColor = true;
            this.OutZoom.Click += new System.EventHandler(this.OutZoom_Click);
            // 
            // InZoom
            // 
            this.InZoom.Image = global::Test.Properties.Resources.放大;
            this.InZoom.Location = new System.Drawing.Point(3, 171);
            this.InZoom.Name = "InZoom";
            this.InZoom.Size = new System.Drawing.Size(50, 50);
            this.InZoom.TabIndex = 6;
            this.toolTip1.SetToolTip(this.InZoom, "放大图片");
            this.InZoom.UseVisualStyleBackColor = true;
            this.InZoom.Click += new System.EventHandler(this.InZoom_Click);
            // 
            // RemoveDay
            // 
            this.RemoveDay.Image = global::Test.Properties.Resources.减少减去减号;
            this.RemoveDay.Location = new System.Drawing.Point(54, 115);
            this.RemoveDay.Name = "RemoveDay";
            this.RemoveDay.Size = new System.Drawing.Size(50, 50);
            this.RemoveDay.TabIndex = 5;
            this.toolTip1.SetToolTip(this.RemoveDay, "减少一天");
            this.RemoveDay.UseVisualStyleBackColor = true;
            this.RemoveDay.Click += new System.EventHandler(this.RemoveDay_Click);
            // 
            // AddDay
            // 
            this.AddDay.Image = global::Test.Properties.Resources.增加添加加号;
            this.AddDay.Location = new System.Drawing.Point(3, 115);
            this.AddDay.Name = "AddDay";
            this.AddDay.Size = new System.Drawing.Size(50, 50);
            this.AddDay.TabIndex = 4;
            this.toolTip1.SetToolTip(this.AddDay, "增加一天");
            this.AddDay.UseVisualStyleBackColor = true;
            this.AddDay.Click += new System.EventHandler(this.AddDay_Click);
            // 
            // MonthRight
            // 
            this.MonthRight.Image = global::Test.Properties.Resources.向右4;
            this.MonthRight.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.MonthRight.Location = new System.Drawing.Point(54, 59);
            this.MonthRight.Name = "MonthRight";
            this.MonthRight.Size = new System.Drawing.Size(50, 50);
            this.MonthRight.TabIndex = 3;
            this.toolTip1.SetToolTip(this.MonthRight, "右移一月");
            this.MonthRight.UseVisualStyleBackColor = true;
            this.MonthRight.Click += new System.EventHandler(this.MonthRight_Click);
            this.MonthRight.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            this.MonthRight.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseMove);
            this.MonthRight.MouseUp += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseUp);
            // 
            // MonthLeft
            // 
            this.MonthLeft.Image = global::Test.Properties.Resources.向左4;
            this.MonthLeft.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.MonthLeft.Location = new System.Drawing.Point(3, 59);
            this.MonthLeft.Name = "MonthLeft";
            this.MonthLeft.Size = new System.Drawing.Size(50, 50);
            this.MonthLeft.TabIndex = 2;
            this.toolTip1.SetToolTip(this.MonthLeft, "左移一月");
            this.MonthLeft.UseVisualStyleBackColor = true;
            this.MonthLeft.Click += new System.EventHandler(this.MonthLeft_Click);
            this.MonthLeft.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            this.MonthLeft.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseMove);
            this.MonthLeft.MouseUp += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseUp);
            // 
            // DayRight
            // 
            this.DayRight.Image = global::Test.Properties.Resources.Right向右;
            this.DayRight.Location = new System.Drawing.Point(54, 3);
            this.DayRight.Name = "DayRight";
            this.DayRight.Size = new System.Drawing.Size(50, 50);
            this.DayRight.TabIndex = 1;
            this.toolTip1.SetToolTip(this.DayRight, "右移一天");
            this.DayRight.UseVisualStyleBackColor = true;
            this.DayRight.Click += new System.EventHandler(this.DayRight_Click);
            this.DayRight.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            this.DayRight.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseMove);
            this.DayRight.MouseUp += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseUp);
            // 
            // DayLeft
            // 
            this.DayLeft.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.DayLeft.Image = global::Test.Properties.Resources.Left向左;
            this.DayLeft.Location = new System.Drawing.Point(3, 3);
            this.DayLeft.Name = "DayLeft";
            this.DayLeft.Size = new System.Drawing.Size(50, 50);
            this.DayLeft.TabIndex = 0;
            this.toolTip1.SetToolTip(this.DayLeft, "左移一天");
            this.DayLeft.UseVisualStyleBackColor = true;
            this.DayLeft.Click += new System.EventHandler(this.DayLeft_Click);
            this.DayLeft.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseDown);
            this.DayLeft.MouseMove += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseMove);
            this.DayLeft.MouseUp += new System.Windows.Forms.MouseEventHandler(this.panel1_MouseUp);
            // 
            // tabControl1
            // 
            this.tabControl1.AutoCloseTabs = true;
            this.tabControl1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.tabControl1.CanReorderTabs = true;
            this.tabControl1.CloseButtonOnTabsAlwaysDisplayed = false;
            this.tabControl1.CloseButtonOnTabsVisible = true;
            this.tabControl1.CloseButtonPosition = DevComponents.DotNetBar.eTabCloseButtonPosition.Right;
            this.tabControl1.CloseButtonVisible = true;
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(3, 3, 3, 100);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedTabFont = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold);
            this.tabControl1.SelectedTabIndex = -1;
            this.tabControl1.Size = new System.Drawing.Size(1032, 800);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.TabLayoutType = DevComponents.DotNetBar.eTabLayoutType.FixedWithNavigationBox;
            this.tabControl1.Tag = "tabControl1";
            this.tabControl1.Text = "tabControl1";
            this.tabControl1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.myFunPictureBox_KeyDown);
            this.tabControl1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.myFunPictureBox_KeyUp);
            // 
            // 画图
            // 
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(1032, 800);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tabControl1);
            this.KeyPreview = true;
            this.Name = "画图";
            this.Text = "电站位置图";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.画图_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.画图_Paint);
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tabControl1)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem SavePicMenuItem;
        private System.Windows.Forms.SaveFileDialog dlgSavePic;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button OutZoom;
        private System.Windows.Forms.Button InZoom;
        private System.Windows.Forms.Button RemoveDay;
        private System.Windows.Forms.Button AddDay;
        private System.Windows.Forms.Button MonthRight;
        private System.Windows.Forms.Button MonthLeft;
        private System.Windows.Forms.ToolTip toolTip1;
        private DevComponents.DotNetBar.TabControl tabControl1;
        private System.Windows.Forms.Button DayRight;
        private System.Windows.Forms.Button DayLeft;
        private System.Windows.Forms.ToolStripMenuItem 隐藏显示图例ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 隐藏显示控制板ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 切换宽高比ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 按住Ctrl鼠标滚轮缩放ToolStripMenuItem;
        //private DevComponents.DotNetBar.TabControlPanel tabControlPanel1;
        //private DevComponents.DotNetBar.TabItem tabItem1;
    }
}

