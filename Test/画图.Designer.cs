
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
            this.保存图片ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
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
            this.保存图片ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(153, 34);
            // 
            // 保存图片ToolStripMenuItem
            // 
            this.保存图片ToolStripMenuItem.Name = "保存图片ToolStripMenuItem";
            this.保存图片ToolStripMenuItem.Size = new System.Drawing.Size(152, 30);
            this.保存图片ToolStripMenuItem.Text = "保存图片";
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
            this.panel1.Location = new System.Drawing.Point(77, 172);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(214, 224);
            this.panel1.TabIndex = 1;
            // 
            // OutZoom
            // 
            this.OutZoom.Location = new System.Drawing.Point(107, 171);
            this.OutZoom.Name = "OutZoom";
            this.OutZoom.Size = new System.Drawing.Size(104, 50);
            this.OutZoom.TabIndex = 7;
            this.OutZoom.Text = "缩小";
            this.toolTip1.SetToolTip(this.OutZoom, "缩小图片");
            this.OutZoom.UseVisualStyleBackColor = true;
            this.OutZoom.Click += new System.EventHandler(this.OutZoom_Click);
            // 
            // InZoom
            // 
            this.InZoom.Location = new System.Drawing.Point(3, 171);
            this.InZoom.Name = "InZoom";
            this.InZoom.Size = new System.Drawing.Size(104, 50);
            this.InZoom.TabIndex = 6;
            this.InZoom.Text = "放大";
            this.toolTip1.SetToolTip(this.InZoom, "放大图片");
            this.InZoom.UseVisualStyleBackColor = true;
            this.InZoom.Click += new System.EventHandler(this.InZoom_Click);
            // 
            // RemoveDay
            // 
            this.RemoveDay.Location = new System.Drawing.Point(107, 115);
            this.RemoveDay.Name = "RemoveDay";
            this.RemoveDay.Size = new System.Drawing.Size(104, 50);
            this.RemoveDay.TabIndex = 5;
            this.RemoveDay.Text = "-";
            this.toolTip1.SetToolTip(this.RemoveDay, "减少一天");
            this.RemoveDay.UseVisualStyleBackColor = true;
            // 
            // AddDay
            // 
            this.AddDay.Location = new System.Drawing.Point(3, 115);
            this.AddDay.Name = "AddDay";
            this.AddDay.Size = new System.Drawing.Size(104, 50);
            this.AddDay.TabIndex = 4;
            this.AddDay.Text = "+";
            this.toolTip1.SetToolTip(this.AddDay, "增加一天");
            this.AddDay.UseVisualStyleBackColor = true;
            // 
            // MonthRight
            // 
            this.MonthRight.Location = new System.Drawing.Point(107, 59);
            this.MonthRight.Name = "MonthRight";
            this.MonthRight.Size = new System.Drawing.Size(104, 50);
            this.MonthRight.TabIndex = 3;
            this.MonthRight.Text = ">>";
            this.toolTip1.SetToolTip(this.MonthRight, "右移一月");
            this.MonthRight.UseVisualStyleBackColor = true;
            // 
            // MonthLeft
            // 
            this.MonthLeft.Location = new System.Drawing.Point(3, 59);
            this.MonthLeft.Name = "MonthLeft";
            this.MonthLeft.Size = new System.Drawing.Size(104, 50);
            this.MonthLeft.TabIndex = 2;
            this.MonthLeft.Text = "<<";
            this.toolTip1.SetToolTip(this.MonthLeft, "左移一月");
            this.MonthLeft.UseVisualStyleBackColor = true;
            // 
            // DayRight
            // 
            this.DayRight.Location = new System.Drawing.Point(107, 3);
            this.DayRight.Name = "DayRight";
            this.DayRight.Size = new System.Drawing.Size(104, 50);
            this.DayRight.TabIndex = 1;
            this.DayRight.Text = ">";
            this.toolTip1.SetToolTip(this.DayRight, "右移一天");
            this.DayRight.UseVisualStyleBackColor = true;
            // 
            // DayLeft
            // 
            this.DayLeft.Location = new System.Drawing.Point(3, 3);
            this.DayLeft.Name = "DayLeft";
            this.DayLeft.Size = new System.Drawing.Size(104, 50);
            this.DayLeft.TabIndex = 0;
            this.DayLeft.Text = "<";
            this.toolTip1.SetToolTip(this.DayLeft, "左移一天");
            this.DayLeft.UseVisualStyleBackColor = true;
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
            this.tabControl1.Size = new System.Drawing.Size(770, 543);
            this.tabControl1.TabIndex = 4;
            this.tabControl1.TabLayoutType = DevComponents.DotNetBar.eTabLayoutType.FixedWithNavigationBox;
            this.tabControl1.Tag = "tabControl1";
            this.tabControl1.Text = "tabControl1";
            // 
            // 画图
            // 
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(1032, 1053);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tabControl1);
            this.Name = "画图";
            this.Text = "电站位置图";
            this.Load += new System.EventHandler(this.画图_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.画图_Paint);
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.tabControl1)).EndInit();
            this.ResumeLayout(false);

        }

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 保存图片ToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button OutZoom;
        private System.Windows.Forms.Button InZoom;
        private System.Windows.Forms.Button RemoveDay;
        private System.Windows.Forms.Button AddDay;
        private System.Windows.Forms.Button MonthRight;
        private System.Windows.Forms.Button MonthLeft;
        private System.Windows.Forms.Button DayRight;
        private System.Windows.Forms.Button DayLeft;
        private System.Windows.Forms.ToolTip toolTip1;
        private DevComponents.DotNetBar.TabControl tabControl1;
        //private DevComponents.DotNetBar.TabControlPanel tabControlPanel1;
        //private DevComponents.DotNetBar.TabItem tabItem1;
    }
}

