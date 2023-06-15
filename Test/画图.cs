﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;

namespace Test
{
    public partial class 画图 : Form
    {
        private MyFunPictureBox myFunPictureBox = null;              //保存图片框控件

        public 画图()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.保存图片ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panel1 = new System.Windows.Forms.Panel();
            this.DayLeft = new System.Windows.Forms.Button();
            this.DayRight = new System.Windows.Forms.Button();
            this.MonthLeft = new System.Windows.Forms.Button();
            this.MonthRight = new System.Windows.Forms.Button();
            this.AddDay = new System.Windows.Forms.Button();
            this.RemoveDay = new System.Windows.Forms.Button();
            this.InZoom = new System.Windows.Forms.Button();
            this.OutZoom = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.contextMenuStrip1.SuspendLayout();
            this.panel1.SuspendLayout();
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
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(214, 224);
            this.panel1.TabIndex = 1;
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
            // 画图
            // 
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(1032, 1053);
            this.Controls.Add(this.panel1);
            this.Name = "画图";
            this.Text = "电站位置图";
            this.Load += new System.EventHandler(this.画图_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.画图_Paint);
            this.contextMenuStrip1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
         
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }

        private void 画图_Load(object sender, EventArgs e)
        {

        }

        private void 画图_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            HatchStyle[] hatchStyles = new HatchStyle[18];
            hatchStyles[0] = HatchStyle.Percent10;
            hatchStyles[1] = HatchStyle.Percent25;
            hatchStyles[2] = HatchStyle.ForwardDiagonal;
            hatchStyles[3] = HatchStyle.BackwardDiagonal;
            hatchStyles[4] = HatchStyle.Sphere;
            hatchStyles[5] = HatchStyle.LightDownwardDiagonal;
            hatchStyles[6] = HatchStyle.LightUpwardDiagonal;
            hatchStyles[7] = HatchStyle.LightVertical;
            hatchStyles[8] = HatchStyle.NarrowVertical;
            hatchStyles[9] = HatchStyle.DiagonalCross;
            hatchStyles[10] = HatchStyle.HorizontalBrick;
            hatchStyles[11] = HatchStyle.OutlinedDiamond;
            hatchStyles[12] = HatchStyle.DiagonalBrick;
            hatchStyles[13] = HatchStyle.Weave;
            hatchStyles[14] = HatchStyle.DarkDownwardDiagonal;
            hatchStyles[15] = HatchStyle.DarkUpwardDiagonal;
            hatchStyles[16] = HatchStyle.DashedDownwardDiagonal;
            hatchStyles[17] = HatchStyle.DashedUpwardDiagonal;
            paintDays(g, 2020, 8, 30, 3, hatchStyles);
        }

        // 开始绘图的年、月、日、绘图天数、一年最大值 
        private void paintDays(Graphics g,int year, int startMonth,int startDay, int days, HatchStyle[] hatchStyle)
        {
            // 左右边界的宽度
            int leftBorderWidth = 80;
            int rightBorderWidth = 15;
            int topBorderWidth = 20;
            int bottomBorderWidth = 220;

            // 整个画面的宽度和高度
            int wholeWidth = 1020;
            int wholeHeight = 950;

            // 首先解析excel文件，获取绘图的信息
            loadData LoadData = new loadData();
            string path = "C:\\power_system\\data\\804\\XML\\input_11_2022020000_70_0.xlsx";
            DataTable STYLdata = loadData.ExcelToDatatable(path, "STYL");
            DataTable NORMdata = loadData.ExcelToDatatable(path, "NORM");
            DataTable Mapsdata = loadData.ExcelToDatatable(path, "MAPs");
            List<Dictionary<string, string>> STYLlist = loadData.STYLTableToData(STYLdata);
            List<Dictionary<string, Dictionary<string, string>>> NORMlist = loadData.NORMTableToData(NORMdata);
            Dictionary<string, Dictionary<string, List<string>>> dictionary = loadData.MAPsTableToData(Mapsdata);
            dictionary = (from d in dictionary orderby d.Key descending select d).ToDictionary(k => k.Key, v => v.Value);

            // 获取最大纵坐标
            Dictionary<string, string> STYLinfo1 = STYLlist[1];
            List<string> temp = new List<string>(STYLinfo1.Keys);
            int maxVal = int.Parse(STYLinfo1[temp[6]]);

            // 绘制坐标轴
            Pen pen = new Pen(Brushes.Black);
            pen.Width = 1.8F;
            // 绘制X轴
            Point start = new Point(leftBorderWidth, wholeHeight - bottomBorderWidth);
            Point end = new Point(wholeWidth - rightBorderWidth, wholeHeight - bottomBorderWidth);
            g.DrawLine(pen, start, end);
            Point[] arrow = new Point[2];
            arrow[0] = new Point(end.X - 5, end.Y + 3);
            arrow[1] = new Point(end.X - 5, end.Y - 3);
            g.DrawLine(pen, arrow[0], end);
            g.DrawLine(pen, arrow[1], end);

            // 绘制Y轴
            Point end2 = new Point(leftBorderWidth, topBorderWidth);
            g.DrawLine(pen, start, end2);
            arrow[0] = new Point(end2.X - 3, end2.Y + 5);
            arrow[1] = new Point(end2.X + 3, end2.Y + 5);
            g.DrawLine(pen, arrow[0], end2);
            g.DrawLine(pen, arrow[1], end2);
            end2.Y += 5;

            // 绘制中间的电站位置图
            

            // 获取到数据后，需要绘图
            // 先得到画图的起始天数
            int startDays = dateToDay(year, startMonth, startDay);
            // 获取数据和颜色、样式来绘图
            // 原始负荷需要额外处理，最后绘制
            Point[] originalLoadPoints = new Point[24 * days * 2];
            foreach (string key in dictionary.Keys)
            {
                string flag = key;
                Dictionary<string, List<string>> dic1 = dictionary[key];
                ArrayList array = new ArrayList();
                for (int i = 0; i < days; i++)
                {
                     string str = Convert.ToString(i + startDays);
                     if (dic1.ContainsKey(str)){
                        List<string> list = dic1[str];
                        foreach(string val in list)
                        {
                            array.Add(int.Parse(val));
                        }  
                     }
                     else
                     {
                        if(array.Count != 0)
                        {
                            for(int j = 0; j < 24; j++)
                            {
                                array.Add(0);
                            }
                        }

                     }
                }
                if (array.Count != 0)
                {
                    int[] data = (int[])array.ToArray(typeof(int));
                    // data-数据 end-绘图的左上角坐标 width-绘图的宽度 height-绘图的高度 maxVal-一年的最大值，用于固定Y轴
                    Point[] points = dataToPoint(data, end2, wholeWidth - leftBorderWidth - rightBorderWidth - 5, wholeHeight - topBorderWidth - bottomBorderWidth - 5, maxVal);
                    // 笔刷中颜色、填充样式应该解析文件得到
                    Dictionary<string, Dictionary<string, string>> NORMdic = NORMlist[0];
                    Dictionary<string, string> NORMinfo = NORMdic[key];
                    int hatchID = int.Parse(NORMinfo["Hatch"]);
                    string color = NORMinfo["ARGB"];
                    Dictionary<string, string> STYLinfo = STYLlist[0];
                    color = STYLinfo[color];
                    string[] ARGBS = color.Split(' ');
                    if (key.Equals("9950"))
                    {
                        originalLoadPoints = points;
                    }
                    else{
                        HatchBrush hBrush = new HatchBrush(hatchStyle[hatchID - 21], Color.Black, Color.FromArgb(int.Parse(ARGBS[0]), int.Parse(ARGBS[1]), int.Parse(ARGBS[2]), int.Parse(ARGBS[3])));
                        // 画图
                        g.FillPolygon(hBrush, points);
                        g.DrawPolygon(Pens.Black, points);
                    }
                }
            }
            Pen blackPen = new Pen(Color.Black, (float)0.5);
            blackPen.DashPattern = new float[] { 5, 4 };
            g.DrawPolygon(blackPen, originalLoadPoints);

            // 绘制完图片部分，再去画坐标轴的间隔
            // 根据天数为1天还是多天，决定X轴分为24小时还是若干天
            int interval;
            if (days > 1)
            {
                interval = (wholeWidth - leftBorderWidth - rightBorderWidth - 5) / (days*24) * 24;
                Point[] downPoints = new Point[days];
                Point[] upPoints = new Point[days];
                String[] dates = getDateString(year, startMonth, startDay, days);
                for (int i = 0; i < days; i++)
                {
                    downPoints[i] = new Point(start.X + interval * i, start.Y);
                    upPoints[i] = new Point(start.X + interval * i, start.Y - 5);
                    g.DrawLine(Pens.Black, downPoints[i], upPoints[i]);
                    Font font = new Font("黑体", 9);
                    g.DrawString(dates[i], font, Brushes.Black, new Point(downPoints[i].X - 8, downPoints[i].Y + 5));
                }
            }
            else
            {
                interval = (end.X - start.X - 5) / 24;
                Point[] downPoints = new Point[24];
                Point[] upPoints = new Point[24];
                for (int i = 0; i < 24; i++)
                {
                    downPoints[i] = new Point(start.X + interval * (i + 1), start.Y);
                    upPoints[i] = new Point(start.X + interval * (i + 1), start.Y - 5);
                    g.DrawLine(Pens.Black, downPoints[i], upPoints[i]);
                    Font font = new Font("黑体", 9);
                    g.DrawString((i + 1) + "时", font, Brushes.Black, new Point(downPoints[i].X - 8, downPoints[i].Y + 5));
                }
            }

            // 绘制Y轴的间隔
            interval = (start.Y - end2.Y - 5) / 10;
            for (int i = 0; i < 10; i++)
            {
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Far; //靠右对齐
                Rectangle space = new Rectangle(start.X - 70, (start.Y - interval * (i + 1)), 60, 10);
                Font font = new Font("黑体", 9);
                g.DrawString((maxVal / 10 * (i + 1)) + "", font, Brushes.Black, space, format);
            }

            // 绘制网格线

        }

        // 数据转换为点坐标
        private Point[] dataToPoint(int[] data, Point start, int width, int height, int maxVal)
        {
            Point[] points = new Point[data.Length * 2 + 2];
            int oneWidth = width / data.Length;
            double oneHeight = (float)height / maxVal;
            for(int i = 0; i < data.Length; i++)
            {
                int left = i * 2;
                int right = left + 1;
                points[left] = new Point(start.X + i * oneWidth, (int)(start.Y + height - data[i] * oneHeight));
                points[right] = new Point(start.X + (i + 1) * oneWidth, (int)(start.Y + height - data[i] * oneHeight));
            }
            points[points.Length - 2] = new Point(points[points.Length - 3].X, start.Y + height);
            points[points.Length - 1] = new Point(points[0].X, start.Y + height);
            return points;
        }
    
        // 根据起始日期和天数，转换为x月x日的字符串
        private String[] getDateString(int year, int month, int day, int days)
        {
            String[] date = new String[days];
            Boolean flag = false;
            if ((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0))
            {
                flag = true;
            }
            int monthDay = 0;
            switch (month)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    monthDay = 31;
                    break;
                case 2:
                    if (flag)
                    {
                        monthDay = 29;
                    }
                    else
                    {
                        monthDay = 28;
                    }
                    break;
                case 4:
                case 6:
                case 9:
                case 11:
                    monthDay = 30;
                    break;
            }
            for(int i = 0; i < days; i++)
            {
                if (day > monthDay)
                {
                    month += 1;
                    day = 1;
                }
                date[i] = month + "月" + day + "日";
                day++;
            }
            return date;
        }

        // 根据x月x日的日期转换为一年之中的天数
        private int dateToDay(int year, int month, int day)
        {
            Boolean flag = false;
            if ((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0))
            {
                flag = true;
            }
            int days = -1;
            for(int i=0; i < month - 1; i++)
            {
                switch (i+1)
                {
                    case 1:
                    case 3:
                    case 5:
                    case 7:
                    case 8:
                    case 10:
                    case 12:
                        days += 31;
                        break;
                    case 2:
                        if (flag)
                        {
                            days += 29;
                        }
                        else
                        {
                            days += 28;
                        }
                        break;
                    case 4:
                    case 6:
                    case 9:
                    case 11:
                        days += 30;
                        break;
                }
            }
            return (days + day);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            // paintDays(g, 2020, 8, 30, 5, 80000);
        }

        private void pictureBox_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            
        }

        // 23.6.15 cjy
        private static void InZoomPic(MyPictureBox picture, Panel panel)
        {
            if (picture.Width * 1.1 <= 1400 && picture.Height * 1.1 <= 2400)
            {
                picture.Size = new Size((int)(picture.Width * 1.1), (int)(picture.Height * 1.1));
                panel.AutoScroll = true;
                picture.smallFontSize = (float)(picture.smallFontSize * 1.1);
                picture.largeFontSize = (float)(picture.largeFontSize * 1.1);
                picture.logoWidth = (int)(picture.logoWidth * 1.1);
                picture.logoPos = new Rectangle((int)(picture.logoPos.Left * 1.1)
                    , (int)(picture.logoPos.Top * 1.1)
                    , picture.logoPos.Width
                    , picture.logoPos.Height); ;
                picture.drawArea = new Rectangle((int)(picture.drawArea.Left * 1.1),
                      (int)(picture.drawArea.Top * 1.1),
                      (int)(picture.drawArea.Width * 1.1),
                      (int)(picture.drawArea.Height * 1.1
                      ));
                picture.Invalidate();
            }
        }

        private static void OutZoomPic(MyPictureBox picture, Panel panel)
        {
            if (picture.Width * 0.9 > 300 && picture.Height * 0.9 > 500)
            {
                picture.Size = new Size((int)(picture.Width * 0.9), (int)(picture.Height * 0.9));
                panel.AutoScroll = true;
                picture.smallFontSize = (float)(picture.smallFontSize * 0.9);
                picture.largeFontSize = (float)(picture.largeFontSize * 0.9);
                picture.logoWidth = (int)(picture.logoWidth * 0.9);
                picture.logoPos = new Rectangle((int)(picture.logoPos.Left * 0.9)
                    , (int)(picture.logoPos.Top * 0.9)
                    , picture.logoPos.Width
                    , picture.logoPos.Height);
                picture.drawArea = new Rectangle((int)(picture.drawArea.Left * 0.9),
                    (int)(picture.drawArea.Top * 0.9),
                    (int)(picture.drawArea.Width * 0.9),
                    (int)(picture.drawArea.Height * 0.9));
                picture.Invalidate();
            }
        }


        private string WrapLogoString(string originalStr)
        {
            int j = 0;
            string result = "";
            for (int i = 0; i < originalStr.Length; i++)
            {
                result += originalStr[i];
                j++;
                if (j == 5)
                {
                    result += "\n";
                    j = 0;
                }
            }
            return result;
        }
        private void SortLogo(MyPictureBox picture)
        {
            for (int i = 0; i < picture.LogoItems.Count; i++)
            {
                int minPriority = 10000;
                int index = -1;
                for (int j = i; j < picture.LogoItems.Count; j++)
                    if (picture.LogoItems[j].priority < minPriority)
                    {
                        minPriority = picture.LogoItems[j].priority;
                        index = j;
                    }
                LogoItem item = new LogoItem();
                item.brush = picture.LogoItems[index].brush;
                item.description = picture.LogoItems[index].description;
                item.priority = picture.LogoItems[index].priority;
                picture.LogoItems.RemoveAt(index);
                picture.LogoItems.Insert(i, item); ;
            }
        }
        //private void DrawLogo(MyPictureBox picture, Graphics g)
        //{
        //    //修改如下 直接将所有的图标画出
        //    //SortLogo(picture);
        //    fillBrushes.Sort();
        //    for (Int32 brushIndex = 0; brushIndex < fillBrushes.Count; brushIndex++)
        //    {
        //        if (fillBrushes[brushIndex].describe.Contains("出力"))
        //        {
        //            LogoItem item = new LogoItem();
        //            item.brush = fillBrushes[brushIndex].myBrush;
        //            item.description = fillBrushes[brushIndex].describe;
        //            item.priority = fillBrushes[brushIndex].priority;
        //            picture.LogoItems.Add(item);
        //        }
        //    }

        //    LogoItem newItem = new LogoItem();
        //    newItem.priority = 0;
        //    newItem.brush = new SolidBrush(Color.SkyBlue);
        //    newItem.description = "原始负荷";
        //    picture.LogoItems.Insert(0, newItem);

        //    //因为吴老师要求将部分图例合并，所以也在此处添加添加by孙凯 2016.1.19
        //    //抽蓄发电/抽蓄填谷合并、电力不足/调峰不足合并、新能源弃电/水电弃水合并
        //    newItem.brush = fillBrushes[myDrawHelper.getBrushArrayIndex(5)].myBrush;
        //    newItem.secondBrush = fillBrushes[myDrawHelper.getBrushArrayIndex(21)].myBrush;
        //    newItem.description = "抽蓄发电/抽蓄填谷";
        //    picture.LogoItems.Insert(1, newItem);

        //    newItem.brush = fillBrushes[myDrawHelper.getBrushArrayIndex(22)].myBrush;
        //    newItem.secondBrush = fillBrushes[myDrawHelper.getBrushArrayIndex(24)].myBrush;
        //    newItem.description = "电力不足/调峰不足";
        //    picture.LogoItems.Insert(2, newItem);

        //    newItem.brush = fillBrushes[myDrawHelper.getBrushArrayIndex(23)].myBrush;
        //    newItem.secondBrush = fillBrushes[myDrawHelper.getBrushArrayIndex(25)].myBrush;
        //    newItem.description = "新能源弃电/水电弃水";
        //    picture.LogoItems.Insert(3, newItem);

        //    //添加Flg = 27、28的Logo
        //    //添加by孙凯 2015.7.7
        //    //Logo 背景使用新能源 此处采用硬编码 为config文件中selectItems的Priority-1
        //    newItem.brush = fillBrushes[myDrawHelper.getBrushArrayIndex(23)].myBrush;
        //    newItem.description = "新能源/风/光发电";
        //    picture.LogoItems.Insert(4, newItem);

        //    //newItem = new LogoItem();
        //    //newItem.priority = 2;
        //    //newItem.brush = new SolidBrush(Color.Green);
        //    //newItem.description = "光伏发电位置曲线";
        //    //picture.LogoItems.Insert(2, newItem);
        //    //添加结束 by 孙凯

        //    SolidBrush backBrush = new SolidBrush(Color.White);

        //    Font drawFont = new Font("宋体", picture.smallFontSize);
        //    SolidBrush drawBrush = new SolidBrush(Color.Black);
        //    Pen framePen = new Pen(Color.Black, 1.0f);
        //    Pen dashPen = new Pen(Color.Black, 1.0f);
        //    dashPen.DashStyle = DashStyle.Dash;
        //    //用于绘制Flg=27的图标
        //    Pen tmpPen1 = new Pen(Color.Orange, 1.0f);
        //    tmpPen1.DashStyle = DashStyle.Dash;
        //    //用于绘制Flg=28的图标
        //    //Pen tmpPen2 = new Pen(Color.HotPink, 1.0f);
        //    //tmpPen2.DashStyle = DashStyle.Dash;

        //    int vacant = 10;
        //    Font titleFont = new Font("宋体", picture.largeFontSize, FontStyle.Bold);

        //    picture.logoPos = new Rectangle(picture.logoPos.Left
        //        , picture.logoPos.Top
        //        , picture.logoWidth
        //        , (picture.LogoItems.Count + 1) / 2 * (drawFont.Height * 2 + vacant) +
        //        vacant + titleFont.Height + vacant + 3);


        //    g.FillRectangle(backBrush, picture.logoPos);
        //    StringFormat stringFormat = new StringFormat();
        //    stringFormat.Alignment = StringAlignment.Center;

        //    int itemWidth = picture.logoPos.Width / 2;

        //    g.DrawString("图例", titleFont, drawBrush, picture.logoPos.Left + picture.logoPos.Width / 2, picture.logoPos.Top + vacant, stringFormat);
        //    Pen pen = new Pen(Color.Black);
        //    g.DrawLine(pen, picture.logoPos.Left, picture.logoPos.Top + vacant + titleFont.Height + 3,
        //        picture.logoPos.Right, picture.logoPos.Top + vacant + titleFont.Height + 3);
        //    Point startPoint = new Point(picture.logoPos.Left, picture.logoPos.Top + vacant + titleFont.Height + 3);


        //    for (int i = 0; i < picture.LogoItems.Count; i++)
        //    {
        //        Point point = new Point(startPoint.X + 5 + (i % 2) * itemWidth,
        //            startPoint.Y + (i / 2) * (drawFont.Height * 2 + vacant) + vacant);
        //        //因为Flg = 0,27,28(对应下标分别为0, 1与其他情况不同故修改
        //        //修改by孙凯 2015.7.7
        //        switch (i)
        //        {
        //            case 0:
        //                g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y + drawFont.Height, 40, drawFont.Height);
        //                g.DrawLine(framePen, point.X, point.Y + drawFont.Height, point.X + 40, point.Y + drawFont.Height);

        //                PointF[] points = new PointF[]
        //                {
        //                    new PointF(point.X, point.Y + (float)drawFont.Height*3.0f/2),
        //                    new PointF(point.X + 20, point.Y + (float)drawFont.Height*3.0f/2),
        //                    new PointF(point.X + 20, point.Y+(float)drawFont.Height/2.0f),
        //                    new PointF(point.X + 40, point.Y+(float)drawFont.Height/2.0f)
        //                };
        //                g.DrawLines(dashPen, points);
        //                g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
        //                break;
        //            //因为吴老师要求将部分图例合并，所以也在此处添加添加by孙凯 2016.1.19
        //            //抽蓄发电/抽蓄填谷合并、电力不足/调峰不足合并、新能源弃电/水电弃水合并
        //            case 1:
        //            case 2:
        //            case 3:
        //                g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y, 40, drawFont.Height);
        //                g.FillRectangle(picture.LogoItems[i].secondBrush, point.X, point.Y + drawFont.Height, 40, drawFont.Height);
        //                g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
        //                break;
        //            case 4:
        //                g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y, 40, drawFont.Height * 2);
        //                //g.DrawLine(framePen, point.X, point.Y + drawFont.Height, point.X + 40, point.Y + drawFont.Height);

        //                PointF[] points3 = new PointF[]
        //                {
        //                    new PointF(point.X, point.Y + (float)drawFont.Height*3.0f/4),
        //                    new PointF(point.X + 20, point.Y + (float)drawFont.Height*3.0f/4),
        //                    new PointF(point.X + 20, point.Y+(float)drawFont.Height/4.0f),
        //                    new PointF(point.X + 40, point.Y+(float)drawFont.Height/4.0f)
        //                };
        //                g.DrawLines(tmpPen1, points3);

        //                PointF[] points4 = new PointF[]
        //                {
        //                    new PointF(point.X, point.Y + (float)drawFont.Height*7.0f/4),
        //                    new PointF(point.X + 12, point.Y + (float)drawFont.Height*7.0f/4),
        //                    new PointF(point.X + 12, point.Y+(float)drawFont.Height*5/4.0f),
        //                    new PointF(point.X + 28, point.Y + (float)drawFont.Height*5.0f/4),
        //                    new PointF(point.X + 28, point.Y+(float)drawFont.Height*7/4.0f),
        //                    new PointF(point.X + 40, point.Y+(float)drawFont.Height*7/4.0f)
        //                };
        //                g.DrawLines(tmpPen1, points4);
        //                g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
        //                break;
        //            //case 2:
        //            //    g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y+drawFont.Height, 40, drawFont.Height );
        //            //    g.DrawLine(framePen, point.X, point.Y + drawFont.Height, point.X + 40, point.Y + drawFont.Height);

        //            //    PointF[] points2 = new PointF[]
        //            //    {
        //            //        new PointF(point.X, point.Y + (float)drawFont.Height*3.0f/2),
        //            //        new PointF(point.X + 20, point.Y + (float)drawFont.Height*3.0f/2),
        //            //        new PointF(point.X + 20, point.Y+(float)drawFont.Height/2.0f),
        //            //        new PointF(point.X + 40, point.Y+(float)drawFont.Height/2.0f)
        //            //    };
        //            //    g.DrawLines(tmpPen2,points2);
        //            //    g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
        //            //    break;
        //            default:
        //                g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y, 40, drawFont.Height * 2);
        //                g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
        //                break;
        //        }
        //        g.DrawRectangle(framePen, point.X, point.Y, 40, drawFont.Height * 2);
        //    }

        //    stringFormat.Dispose();
        //    pen.Dispose();
        //    framePen.Dispose();
        //    drawBrush.Dispose();
        //    titleFont.Dispose();
        //    drawFont.Dispose();
        //}


        private void InZoom_Click(object sender, EventArgs e)
        {
            //InZoomPic(myFunPictureBox, myFunPictureBox.Parent as Panel);
        }

        private void OutZoom_Click(object sender, EventArgs e)
        {
            //OutZoomPic(myFunPictureBox, myFunPictureBox.Parent as Panel);
        }
    }

    public class loadData
    {
        public static DataTable ExcelToDatatable(string fileName, string sheetName)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 1;//第一行均为标题
            IWorkbook workbook = null;
            FileStream fs;
            int cellCount = 0; //需要处理数据的列数
            int rowCount = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                {
                    workbook = new XSSFWorkbook(fs);
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                {
                    workbook = new HSSFWorkbook(fs);
                }
                if(sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                }
                else
                {
                    Console.WriteLine("SheetName is null");
                }
                if(sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    if (sheetName.Equals("STYL"))
                    {
                        cellCount = 3;
                    }else if (sheetName.Equals("NORM"))
                    {
                        cellCount = 6;
                    }else if (sheetName.Equals("MAPs"))
                    {
                        cellCount = 26;
                    }
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)//第一行列数循环
                    {
                        DataColumn column = new DataColumn(firstRow.GetCell(i).StringCellValue);//获取标题
                        data.Columns.Add(column);//添加列
                    }
                }
                else
                {
                    Console.WriteLine("Sheet is null!");
                }
                rowCount = sheet.LastRowNum;
                for (int i = startRow; i <= rowCount; ++i)//循环遍历所有行
                {
                    IRow row = sheet.GetRow(i);//第几行
                    int j = row.FirstCellNum;
                    if(j < 0)
                    {
                        break;
                    }
                    //将excel表每一行的数据添加到datatable的行中
                    DataRow dataRow = data.NewRow();
                    for (; j < cellCount; ++j)
                    {
                        if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                        {
                            dataRow[j] = row.GetCell(j).ToString();
                        }
                    }
                    data.Rows.Add(dataRow);
                }
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        public static List<Dictionary<string,string>> STYLTableToData(DataTable STYLDt)
        {
            List <Dictionary<string, string>> STYLData = new List<Dictionary<string, string>>();
            Dictionary<string, string> colorDictionary = new Dictionary<string, string>();
            Dictionary<string, string> drawDictionary = new Dictionary<string, string>();
            STYLData.Add(colorDictionary);
            STYLData.Add(drawDictionary);
            int startRow = 0;
            string data = STYLDt.Rows[startRow][1].ToString();
            while (!(data.IndexOf("绘图") > 0))
            {
                if(data.IndexOf("、") > 0)
                {
                    startRow++;
                    data = STYLDt.Rows[startRow][1].ToString();
                    continue;
                }
                else
                {
                    colorDictionary.Add(STYLDt.Rows[startRow]["ID"].ToString(), data);
                    startRow++;
                    data = STYLDt.Rows[startRow]["Item"].ToString();
                }    
            }
            startRow++;
            for(;startRow < STYLDt.Rows.Count; startRow++)
            {
                drawDictionary.Add(STYLDt.Rows[startRow]["备注"].ToString(), STYLDt.Rows[startRow]["Item"].ToString());
            }
            return STYLData;
        }

        public static List<Dictionary<string, Dictionary<string, string>>> NORMTableToData(DataTable NORMDt)
        {
            List<Dictionary<string, Dictionary<string, string>>> NORMData = new List<Dictionary<string, Dictionary<string, string>>>();
            Dictionary<string, Dictionary<string, string>> drawParameter= new Dictionary<string, Dictionary<string, string>>();
            NORMData.Add(drawParameter);
            int startRow = 0;
            while (!(NORMDt.Rows[startRow]["Item"].ToString().IndexOf("绘图参数") > 0))
            {
                startRow++;
            }
            startRow++;
            for(; startRow < NORMDt.Rows.Count; startRow++)
            {
                Dictionary<string, string> dictionary = new Dictionary<string, string>();
                dictionary.Add("Hatch", NORMDt.Rows[startRow]["Hatch"].ToString());
                dictionary.Add("ARGB", NORMDt.Rows[startRow]["ARGB"].ToString());
                dictionary.Add("Mark", "1");
                if (!(drawParameter.ContainsKey(NORMDt.Rows[startRow]["Flag"].ToString())))
                {
                    drawParameter.Add(NORMDt.Rows[startRow]["Flag"].ToString(), dictionary);
                }
                
            }
            return NORMData;
        }
        public static Dictionary<string, Dictionary<string, List<string>>> MAPsTableToData(DataTable MAPsDt)
        {
            Dictionary<string, Dictionary<string, List<string>>> flagDictionary = new Dictionary<string, Dictionary<string, List<string>>>();
            int startRow = 0;
            for(;startRow < MAPsDt.Rows.Count; startRow++)
            {
                string flag = MAPsDt.Rows[startRow]["Flag"].ToString();
                Dictionary<string, List<string>> dateDictionary;
                List<string> data = new List<string>();
                if (!flagDictionary.ContainsKey(flag))
                {
                    dateDictionary = new Dictionary<string, List<string>>();
                    flagDictionary.Add(flag, dateDictionary);
                }
                else
                {
                    flagDictionary.TryGetValue(flag, out dateDictionary);
                }
                for(int startLine = 1; startLine <= 24; startLine++)
                {
                    data.Add(MAPsDt.Rows[startRow][startLine].ToString());
                }
                dateDictionary.Add(MAPsDt.Rows[startRow]["Date"].ToString(), data);
            }
            return flagDictionary;
        }
    }

    public class MyFunPictureBox : MyPictureBox
    {
        new public List<DataTable> LevelLines { get; set; }
    }

    public struct LogoItem
    {
        public Brush brush;
        //因为一些图标合并到一起所以需要保存两个Brush 添加by孙凯 2016.1.19
        public Brush secondBrush;
        public string description;
        public int priority;
    }
    public class MyPictureBox : PictureBox
    {
        public DataTable LevelLines { get; set; }
        public DataTable genPos { get; set; }
        public Rectangle logoPos { get; set; }
        public List<LogoItem> LogoItems = new List<LogoItem>();
        public Point previousPos { get; set; }
        public float largeFontSize = 12;
        public float smallFontSize = 10.5F;
        public System.Drawing.Printing.PageSettings pageSettings { get; set; }
        public Rectangle drawArea { get; set; }
        public int logoWidth = 300;
        public bool drawed = false;
        public int maxRectangleY = 10000;
    }
}
