using System;
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
using DevComponents.DotNetBar;
using System.Drawing.Printing;


using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections;

namespace Test
{
    public partial class 画图 : Form
    {
        private MyFunPictureBox myFunPictureBox = null;              //保存图片框控件
        private List<Dictionary<string, string>> STYLlist;
        private List<Dictionary<string, Dictionary<string, string>>> NORMlist;
        private Dictionary<string, Dictionary<string, List<string>>> dictionary;

        public 画图()
        {
            InitializeComponent();
            string path = "C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-06\\input_11_2022020000_70_0.xlsx";
            //string path = "C:\\power_system\\data\\804\\XML\\input_11_2022020000_70_0.xlsx";
            loadData LoadData = new loadData();
            DataTable STYLdata = loadData.ExcelToDatatable(path, "STYL");
            DataTable NORMdata = loadData.ExcelToDatatable(path, "NORM");
            DataTable Mapsdata = loadData.ExcelToDatatable(path, "MAPs");
            this.STYLlist = loadData.STYLTableToData(STYLdata);
            this.NORMlist = loadData.NORMTableToData(NORMdata);
            this.dictionary = loadData.MAPsTableToData(Mapsdata);
            this.dictionary = (from d in dictionary orderby d.Key descending select d).ToDictionary(k => k.Key, v => v.Value);
        }

        public void newTab()
        {
            TabItem tp  = this.tabControl1.CreateTab("图1");

            TabControlPanel tcp = new TabControlPanel();
            tcp.Visible = false;
            tcp.TabItem = tp;
            tcp.AutoScroll = true;
            tcp.Dock = DockStyle.Fill;

            tcp.Resize += new System.EventHandler(this.pictureBox_Resize);

            myFunPictureBox = new MyFunPictureBox();

            tcp.Controls.Add(myFunPictureBox);


            myFunPictureBox.AutoScrollOffset = new Point(0, 0);

            tcp.ScrollControlIntoView(myFunPictureBox);

            //属性设置
            //myFunPictureBox.LevelLines = d;
            //TODO 注释by孙凯 2015.12.17 pictureBox.genPos=d[2*dtIndex+1];

            myFunPictureBox.pageSettings = new System.Drawing.Printing.PageSettings();
            PageSettings pageSettings = myFunPictureBox.pageSettings;
            //TODO 暂时修改by孙凯 2016.1.6

            myFunPictureBox.Width = (int)(pageSettings.PaperSize.Width / 100.0 * 96);
            myFunPictureBox.Height = (int)(pageSettings.PaperSize.Height / 100.0 * 96);

            int xDraw = (int)(pageSettings.Margins.Left / 254.0 * 96);
            int yDraw = (int)(pageSettings.Margins.Top / 254.0 * 96);
            int widthDraw = (int)(pageSettings.PaperSize.Width / 100.0 * 96) -
                (int)((pageSettings.Margins.Left + pageSettings.Margins.Right) / 254.0 * 96);
            int heightDraw = (int)(pageSettings.PaperSize.Height / 100.0 * 96) -
                (int)((pageSettings.Margins.Top + pageSettings.Margins.Bottom) / 254.0 * 96);
            myFunPictureBox.drawArea = new Rectangle(xDraw, yDraw, widthDraw, heightDraw);

            myFunPictureBox.logoPos = new Rectangle(myFunPictureBox.drawArea.Left + (int)(myFunPictureBox.drawArea.Width * 0.5)
                , myFunPictureBox.drawArea.Bottom - (int)(myFunPictureBox.drawArea.Height * 0.4)
                , 300
                , (myFunPictureBox.LogoItems.Count + 1) / 2 * 50);



            myFunPictureBox.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseWheel);
            myFunPictureBox.MouseLeave += new System.EventHandler(this.pictureBox_MouseLeave);
            myFunPictureBox.MouseEnter += new System.EventHandler(this.pictureBox_MouseEnter);
            myFunPictureBox.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseMove);
            myFunPictureBox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseDown);
            myFunPictureBox.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseUp);
            myFunPictureBox.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBox_Paint);
            myFunPictureBox.Resize += new System.EventHandler(this.pictureBox_Resize);


            this.tabControl1.Controls.Add(tcp);
            tp.AttachedControl = tcp;
        }

        private void pictureBox_Resize(object sender, EventArgs e)
        {
            Control picture;
            Control parent;
            if (sender is PictureBox || sender is Panel)
            {
                if (sender is PictureBox)
                {
                    picture = sender as Control;
                    parent = picture.Parent;
                }
                else
                {
                    parent = sender as Control;
                    picture = parent.Controls[0];
                }
                int x = picture.Width < parent.Width ? (parent.Width - picture.Width) / 2 : 0;
                int y = picture.Height < parent.Height ? (parent.Height - picture.Height) / 2 : 0;
                picture.Location = new Point(x, y);

            }
        }

        private void pictureBox_MouseEnter(object sender, System.EventArgs e)
        {
            //((sender as PictureBox).Parent as Panel).Focus();
            (sender as PictureBox).Focus();
        }

        private void pictureBox_MouseLeave(object sender, System.EventArgs e)
        {
            // this.Focus();
        }

        //指示是否正在进行图例拖动
        private bool isDragPic = false;

        private void pictureBox_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            MyPictureBox picture = sender as MyPictureBox;
            picture.previousPos = new Point(e.X, e.Y);

            //指示开始图例拖动
            if (e.Button == MouseButtons.Left && picture.logoPos.Contains(new Point(e.X, e.Y)))
                this.isDragPic = true;
        }

        private void pictureBox_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            MyPictureBox picture = sender as MyPictureBox;

            //指示结束图例拖动
            if (e.Button == MouseButtons.Left)
                this.isDragPic = false;
        }




        bool ctrl = false;
        private void tabControl1_KeyDown(object sender, KeyEventArgs e)
        {
            ctrl = e.Control;
        }

        private void tabControl1_KeyUp(object sender, KeyEventArgs e)
        {
            ctrl = e.Control;
        }

        private void pictureBox_MouseWheel(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (ctrl)
            {
                MyPictureBox picture = sender as MyPictureBox;
                Panel panel = picture.Parent as Panel;
                panel.AutoScroll = false;
                if (e.Delta < 0)
                {
                    shrinkPicture(picture, panel);
                    panel.AutoScroll = true;
                }
                else
                {
                    amplifyPicture(picture, panel);
                }
            }
            else
            {
                Point mousePoint = new Point(e.X, e.Y);
                Panel panel = (sender as PictureBox).Parent as Panel;
                mousePoint.Offset(this.Location.X, this.Location.Y);
                if (panel.RectangleToScreen(panel.DisplayRectangle).Contains(mousePoint))
                {
                    panel.AutoScrollPosition = new Point(0, panel.VerticalScroll.Value - e.Delta);
                }
            }
        }

        private void pictureBox_Paint(object sender, PaintEventArgs e)
        {
            MyFunPictureBox picture = sender as MyFunPictureBox;

            try
            {
                Graphics g = e.Graphics; ;
                g.Clear(Color.White);

                picture.LogoItems.Clear();

                //画图_Paint(picture, e);


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

                //myDrawHelper.drawAxes(picture, g);

                //DrawLogo(picture, g);

                picture.drawed = true;
                // picture.Image = memImage;
            }
            catch (Exception ex)
            {
                //ex.WriteLog();
                MessageBox.Show(ex.ToString());
            }

        }

        private static void amplifyPicture(MyPictureBox picture, Panel panel)
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

        private static void shrinkPicture(MyPictureBox picture, Panel panel)
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


        private void button1_Click(object sender, EventArgs e)
        {

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
            foreach (string key in this.dictionary.Keys)
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
            if (e.Button == MouseButtons.Left)
            {
                MyPictureBox picture = sender as MyPictureBox;
                Point mousePoint = new Point(e.X, e.Y);
                if (isDragPic == true) //2013-9-22 刘水兵：当鼠标拖离图例的时候，应该也是要进行移动的.所以 用一个变量指示 是否正在拖动
                {//if(picture.logoPos.Countains(mousePoint))
                    Point topleft = new Point(picture.logoPos.Left, picture.logoPos.Top);
                    picture.logoPos = new Rectangle(
                        picture.logoPos.Left + e.X - picture.previousPos.X,
                        picture.logoPos.Top + e.Y - picture.previousPos.Y,
                        picture.logoPos.Width,
                        picture.logoPos.Height);
                    picture.previousPos = mousePoint;
                    picture.Invalidate(new Rectangle(topleft.X, topleft.Y,
                        picture.Width + e.X - picture.previousPos.X, picture.logoPos.Height + e.Y - picture.previousPos.Y));
                }
            }
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
            InZoomPic(myFunPictureBox, myFunPictureBox.Parent as Panel);
        }

        private void OutZoom_Click(object sender, EventArgs e)
        {
            OutZoomPic(myFunPictureBox, myFunPictureBox.Parent as Panel);
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
