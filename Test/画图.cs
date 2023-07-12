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
        //private DrawHelper drawHelper;
        private Int32 currentDay;
        private Int32 currentMonth;
        private Int32 currentYear;
        private Int32 currentSpan;

        private Boolean isPanelDrag = false;
        private Point panelPrePosition;

        public 画图()
        {
            InitializeComponent();
            //string path = "C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-06\\input_11_2022020000_70_0.xlsx";
            string path = "C:\\power_system\\data\\804\\XML\\input_11_2022020000_70_0.xlsx";
            loadData LoadData = new loadData();
            DataTable STYLdata = loadData.ExcelToDatatable(path, "STYL");
            DataTable NORMdata = loadData.ExcelToDatatable(path, "NORM");
            DataTable Mapsdata = loadData.ExcelToDatatable(path, "MAPs");
            STYLlist = loadData.STYLTableToData(STYLdata);
            NORMlist = loadData.NORMTableToData(NORMdata);
            dictionary = loadData.MAPsTableToData(Mapsdata);
            dictionary = (from d in dictionary orderby d.Key descending select d).ToDictionary(k => k.Key, v => v.Value);
        }

        private void 画图_Load(object sender, EventArgs e)
        {
            //tabControl1.Focus();
            //Application.Idle += new EventHandler(Application_Idle);
        }
        //void Application_Idle(object sender, EventArgs e)
        //{
        //    // 取消订阅，确保只运行一次
        //    Application.Idle -= new EventHandler(Application_Idle);

        //    // 设置TabControl为焦点
        //    tabControl1.Focus();
        //}

        private void 画图_Paint(object sender, PaintEventArgs e)
        {

        }

        public void newTab(int startDay, int startMonth, int startYear)
        {
            TabItem tp = this.tabControl1.CreateTab("图1");

            TabControlPanel tcp = new TabControlPanel();
            tcp.Visible = false;
            tcp.TabItem = tp;
            tcp.AutoScroll = true;
            tcp.Dock = DockStyle.Fill;

            tcp.Resize += new System.EventHandler(this.pictureBox_Resize);

            //this.drawHelper = new DrawHelper(startYear, startMonth, startDay, 3);
            currentDay = startDay;
            currentMonth = startMonth;
            currentYear = startYear;
            currentSpan = 4;

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

            myFunPictureBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.myFunPictureBox_KeyDown);
            myFunPictureBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.myFunPictureBox_KeyUp);

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
            //(sender as PictureBox).Focus();
            myFunPictureBox.Focus();
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
        private void myFunPictureBox_KeyDown(object sender, KeyEventArgs e)
        {
            ctrl = e.Control;
        }

        private void myFunPictureBox_KeyUp(object sender, KeyEventArgs e)
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
                    OutZoomPic(picture, panel);
                    panel.AutoScroll = true;
                }
                else
                {
                    InZoomPic(picture, panel);
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
                Graphics g = e.Graphics;
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

                paintDays(g, picture, currentYear, currentMonth, currentDay, currentSpan, hatchStyles);

                //myDrawHelper.drawAxes(picture, g);

                DrawLogo(picture, g, hatchStyles);

                picture.drawed = true;
                // picture.Image = memImage;
            }
            catch (Exception ex)
            {
                //ex.WriteLog();
                MessageBox.Show(ex.ToString());
            }

        }

        // 开始绘图的年、月、日、绘图天数、一年最大值 
        private void paintDays(Graphics g, MyFunPictureBox picture, int year, int startMonth, int startDay, int days, HatchStyle[] hatchStyle)
        {
            //Point zeroPoint = new Point(picture.drawArea.Left + (int)(picture.drawArea.Width * 0.1), picture.drawArea.Top + (int)(picture.drawArea.Height * 0.9));
            //setXYInterval(this.myFunPictureBox.drawArea.Width, this.myFunPictureBox.drawArea.Height);
            //myDrawHelper.setZeroPoint(zeroPoint.X, zeroPoint.Y);
            // 左右边界的宽度
            int leftBorderWidth = (int)(picture.drawArea.Width * 0.1);
            int rightBorderWidth = (int)(picture.drawArea.Width * 0.01);
            int topBorderWidth = (int)(picture.drawArea.Height * 0.05);
            int bottomBorderWidth = (int)(picture.drawArea.Height * 0.01);

            // 整个画面的宽度和高度
            int wholeWidth = this.myFunPictureBox.drawArea.Width;
            int wholeHeight = this.myFunPictureBox.drawArea.Height;

            int width = wholeWidth - leftBorderWidth - rightBorderWidth - 10;

            float fontSize = width / 70;//根据图片的绘制宽度调整字体的大小
            if (fontSize > 14)
            {
                fontSize = 14;
            }
            else if (fontSize < 7)
            {
                fontSize = 7;
            }
            int hours = 24 * days;
            int hourLeft = width % hours;//由于整数存在缩进，需要对不能进行整除的部分进行处理，选择的方式是将多余的部分平均分配到每一天的前i个小时中
            int dayRemainder = hourLeft % days;
            int prehours = hourLeft / days;
            int hourWidth = width / hours;

            int[] coordinateList = new int[hours + 1];

            coordinateList[0] = 0;

            for (int i = 1; i <= hours; i++)//将不能平均分配的长度分别分配到每一天中，其中dayRemainder参数说明前dayRemainder相较于其余天宽度长1px，prehours说明每天的前prehours小时的宽度长1px
            {
                int day = i / 24;
                int hour = i % 24;
                coordinateList[i] = coordinateList[i - 1] + hourWidth;

                if (day < dayRemainder)
                {
                    if (hour < prehours + 1)
                    {
                        coordinateList[i]++;
                    }
                }
                else if (hour < prehours)
                {
                    coordinateList[i]++;
                }
            }




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

            //int task = 0;

            foreach (string key in this.dictionary.Keys)
            {
                string flag = key;
                Dictionary<string, List<string>> dic1 = dictionary[key];
                ArrayList array = new ArrayList();
                for (int i = 0; i < days; i++)
                {
                    string str = Convert.ToString(i + startDays);
                    if (dic1.ContainsKey(str))
                    {
                        List<string> list = dic1[str];
                        foreach (string val in list)
                        {
                            array.Add(int.Parse(val));
                        }
                    }
                    else
                    {
                        if (array.Count != 0)
                        {
                            for (int j = 0; j < 24; j++)
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
                    //Console.WriteLine(days + ":" + (wholeWidth - leftBorderWidth - rightBorderWidth - 5));

                    int height = wholeHeight - topBorderWidth - bottomBorderWidth - 5;


                    Point[] points = dataToPoint(data, end2, coordinateList, height, maxVal);
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
                    else
                    {
                        HatchBrush hBrush = new HatchBrush(hatchStyle[hatchID - 21], Color.Black, Color.FromArgb(int.Parse(ARGBS[0]), int.Parse(ARGBS[1]), int.Parse(ARGBS[2]), int.Parse(ARGBS[3])));
                        // 画图
                        g.FillPolygon(hBrush, points);
                        g.DrawPolygon(Pens.Black, points);
                    }
                }
                //task++;
                //if(task > 1)
                //{
                //break;
                //}
            }
            Pen blackPen = new Pen(Color.Black, (float)0.5);
            blackPen.DashPattern = new float[] { 5, 4 };
            g.DrawPolygon(blackPen, originalLoadPoints);


            Point top;
            Pen dotted = new Pen(Color.Black, 1);
            dotted.DashStyle = System.Drawing.Drawing2D.DashStyle.Custom;
            dotted.DashPattern = new float[] { 1, 1 };

            // 绘制完图片部分，再去画坐标轴的间隔
            // 根据天数为1天还是多天，决定X轴分为24小时还是若干天
            int interval;
            if (days > 1)
            {
                //interval = (wholeWidth - leftBorderWidth - rightBorderWidth - 5) / (days*24) * 24;
                Point[] downPoints = new Point[days + 1];
                Point[] upPoints = new Point[days];

                String[] dates = getDateString(year, startMonth, startDay, days);

                for (int i = 0; i < days; i++)
                {
                    downPoints[i] = new Point(start.X + coordinateList[i * 24], start.Y);
                    upPoints[i] = new Point(start.X + coordinateList[i * 24], start.Y - 5);
                    top = new Point(start.X + coordinateList[i * 24], end2.Y + 8);
                    g.DrawLine(Pens.Black, downPoints[i], upPoints[i]);
                    g.DrawLine(dotted, downPoints[i], top);
                    Font font = new Font("黑体", fontSize);
                    g.DrawString(dates[i], font, Brushes.Black, new Point(downPoints[i].X - 8, downPoints[i].Y + 5));
                }
                downPoints[days] = new Point(start.X + coordinateList[days * 24], start.Y);
                top = new Point(start.X + coordinateList[days * 24], end2.Y + 8);
                g.DrawLine(dotted, downPoints[days], top);

            }
            else
            {
                String[] dates = getDateString(year, startMonth, startDay, days);
                Font font = new Font("黑体", fontSize);
                g.DrawString(dates[0], font, Brushes.Black, new Point(start.X - 28, start.Y + 5));
                //interval = (end.X - start.X - 5) / 24;
                Point[] downPoints = new Point[24];
                Point[] upPoints = new Point[24];
                for (int i = 0; i < 24; i++)
                {

                    downPoints[i] = new Point(start.X + coordinateList[i + 1], start.Y);
                    upPoints[i] = new Point(start.X + coordinateList[i + 1], start.Y - 5);
                    top = new Point(start.X + coordinateList[i + 1], end2.Y + 8);
                    g.DrawLine(Pens.Black, downPoints[i], upPoints[i]);
                    g.DrawLine(dotted, downPoints[i], top);
                    g.DrawString((i + 1) + "时", font, Brushes.Black, new Point(downPoints[i].X - 8, downPoints[i].Y + 5));
                }
            }

            // 绘制Y轴的间隔
            interval = (start.Y - end2.Y - 5) / 10;
            for (int i = 0; i < 10; i++)
            {
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Far; //靠右对齐
                Rectangle space = new Rectangle(start.X - 70, (start.Y - interval * (i + 1)), 60, 20);
                Point leftPoint = new Point(start.X, start.Y - interval * (i + 1));
                Point rightPoint = new Point(start.X + 5, start.Y - interval * (i + 1));
                top = new Point(end.X - 10, start.Y - interval * (i + 1));
                g.DrawLine(Pens.Black, leftPoint, rightPoint);
                g.DrawLine(dotted, leftPoint, top);
                Font font = new Font("黑体", fontSize);
                g.DrawString((maxVal / 10 * (i + 1)) + "", font, Brushes.Black, space, format);
            }

            // 绘制网格线

        }

        // 数据转换为点坐标
        private Point[] dataToPoint(int[] data, Point start, int[] coordinateList, int height, int maxVal)
        {
            //Console.WriteLine(width);
            Point[] points = new Point[data.Length * 2 + 2];
            int days = data.Length / 24;

            int X = start.X;
            double oneHeight = (float)height / maxVal;//每1单位对应的纵坐标长度
            int pointIndex = 0;
            for (int i = 0; i < days; i++)
            {
                for (int j = 0; j < 24; j++)//按照每天24小时进行作图
                {
                    int hours = i * 24 + j;//当前坐标点所表示的总小时位置
                    points[pointIndex] = new Point(X + coordinateList[hours], (int)(start.Y + height - data[hours] * oneHeight));
                    pointIndex++;
                    points[pointIndex] = new Point(X + coordinateList[hours + 1], (int)(start.Y + height - data[hours] * oneHeight));
                    pointIndex++;
                }

            }
            //int X = 0;
            //int oneWidth = width / data.Length;
            //for(int i = 0; i < data.Length; i++)
            //{

            //int left = i * 2;
            //int right = left + 1;

            //}
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
            for (int i = 0; i < days; i++)
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
            for (int i = 0; i < month - 1; i++)
            {
                switch (i + 1)
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
        private void DrawLogo(MyPictureBox picture, Graphics g, HatchStyle[] hatchStyles)
        {
            LogoItem newItem = new LogoItem();
            newItem.brush = new SolidBrush(Color.SkyBlue);
            newItem.description = "原始负荷";
            picture.LogoItems.Add(newItem);
            //    newItem.priority = 0;

            Dictionary<String, Dictionary<String, String>> LogoParameter = NORMlist[1];
            Dictionary<string, string> STYLinfo = STYLlist[0];


            for (int i = 4; i <= 20; i++)
            {
                Dictionary<String, String> Info = LogoParameter[i.ToString()];
                int firstHatchID = int.Parse(Info["firstHatch"]);
                int secondHatchID = int.Parse(Info["secondHatch"]);
                string firstColor = Info["firstARGB"];
                string secondColor = Info["secondARGB"];
                string firstItem = Info["firstItem"];
                string secondItem = Info["secondItem"];
                //从Info中取出ARGB编号，再从STYLlist[0]中取出编号对应的ARGB


                Brush firstBrush = infoToBrush(firstColor, firstHatchID);
                Brush secondBrush = infoToBrush(secondColor, secondHatchID);

                newItem = new LogoItem();
                newItem.brush = firstBrush;
                newItem.secondBrush = secondBrush;
                if (secondItem == "")
                {
                    if (firstItem == "")
                    {
                        continue;
                    }
                    newItem.description = firstItem;
                }
                else
                {
                    newItem.description = firstItem + "/" + secondItem;
                }
                picture.LogoItems.Add(newItem);
            }


            SolidBrush backBrush = new SolidBrush(Color.White);
            Font drawFont = new Font("宋体", picture.smallFontSize);

            SolidBrush drawBrush = new SolidBrush(Color.Black);
            Pen framePen = new Pen(Color.Black, 1.0f);
            Pen dashPen = new Pen(Color.Black, 1.0f);
            dashPen.DashStyle = DashStyle.Dash;

            int vacant = 10;
            Font titleFont = new Font("宋体", picture.largeFontSize, FontStyle.Bold);


            picture.logoPos = new Rectangle(picture.logoPos.Left
                , picture.logoPos.Top
                , picture.logoWidth
                , (picture.LogoItems.Count + 1) / 2 * (drawFont.Height * 2 + vacant) +
                vacant + titleFont.Height + vacant + 3);

            g.FillRectangle(backBrush, picture.logoPos);
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;

            int itemWidth = picture.logoPos.Width / 2;

            g.DrawString("图例", titleFont, drawBrush, picture.logoPos.Left + picture.logoPos.Width / 2, picture.logoPos.Top + vacant, stringFormat);
            Pen pen = new Pen(Color.Black);
            g.DrawLine(pen, picture.logoPos.Left, picture.logoPos.Top + vacant + titleFont.Height + 3,
                    picture.logoPos.Right, picture.logoPos.Top + vacant + titleFont.Height + 3);
            Point startPoint = new Point(picture.logoPos.Left, picture.logoPos.Top + vacant + titleFont.Height + 3);

            for (int i = 0; i < picture.LogoItems.Count; i++)
            {
                Point point = new Point(startPoint.X + 5 + (i % 2) * itemWidth,
                    startPoint.Y + (i / 2) * (drawFont.Height * 2 + vacant) + vacant);
                if (i == 0)
                {
                    g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y + drawFont.Height, 40, drawFont.Height);
                    g.DrawLine(framePen, point.X, point.Y + drawFont.Height, point.X + 40, point.Y + drawFont.Height);

                    PointF[] points = new PointF[]
                    {
                                new PointF(point.X, point.Y + (float)drawFont.Height*3.0f/2),
                                new PointF(point.X + 20, point.Y + (float)drawFont.Height*3.0f/2),
                                new PointF(point.X + 20, point.Y+(float)drawFont.Height/2.0f),
                                new PointF(point.X + 40, point.Y+(float)drawFont.Height/2.0f)
                        };
                    g.DrawLines(dashPen, points);
                    g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
                }
                else
                {
                    g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y, 40, drawFont.Height);
                    g.FillRectangle(picture.LogoItems[i].secondBrush, point.X, point.Y + drawFont.Height, 40, drawFont.Height);
                    g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 45, point.Y);
                }
                g.DrawRectangle(framePen, point.X, point.Y, 40, drawFont.Height * 2);
            }

            stringFormat.Dispose();
            pen.Dispose();
            framePen.Dispose();
            drawBrush.Dispose();
            titleFont.Dispose();
            drawFont.Dispose();
        }
        
        //通过NORM表格中的color和hatch的编码，在Style表格中获取对应的填充风格和颜色类型，并生成对应的笔刷
        private Brush infoToBrush(string colorID, int hatchID)
        {
            Dictionary<string, string> STYLinfo = STYLlist[0];

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
            
            //当hatchID=39，也就是超出了hatchStyles的列表范围时，代表的是无填充风格，因此生成只包含颜色的SolidBrush笔刷，否则生成带有填充风格的HatchBrush笔刷
            if (hatchID < hatchStyles.Length + 21)
            {          
                string color = STYLinfo[colorID];
                string[] ARGBS = color.Split(' ');
                HatchBrush hBrush = new HatchBrush(hatchStyles[hatchID - 21], Color.Black, Color.FromArgb(int.Parse(ARGBS[0]), int.Parse(ARGBS[1]), int.Parse(ARGBS[2]), int.Parse(ARGBS[3])));
                return hBrush;
            }
            else
            {
                string color = STYLinfo[colorID];
                string[] ARGBS = color.Split(' ');
                SolidBrush sBrush = new SolidBrush(Color.FromArgb(int.Parse(ARGBS[0]), int.Parse(ARGBS[1]), int.Parse(ARGBS[2]), int.Parse(ARGBS[3])));
                return sBrush;
            }
        }






        private void InZoom_Click(object sender, EventArgs e)
        {
            InZoomPic(myFunPictureBox, myFunPictureBox.Parent as Panel);
        }

        private void OutZoom_Click(object sender, EventArgs e)
        {
            OutZoomPic(myFunPictureBox, myFunPictureBox.Parent as Panel);
        }

        private void DayLeft_Click(object sender, EventArgs e)
        {
            Boolean flag = true;
            if (currentDay > 1)
                currentDay--;
            else if (currentDay == 1 && currentMonth == 1)
                flag = false;
            else
            {
                currentMonth--;
                currentDay = getMaxDay(currentYear, currentMonth);
            }

            if (flag)
                this.myFunPictureBox.Invalidate();
        }

        private void DayRight_Click(object sender, EventArgs e)
        {
            Boolean flag = true;
            Int32 maxDay = getMaxDay(currentYear, currentMonth);
            if (currentDay < maxDay)
                currentDay++;
            else if (currentMonth < 12)
            {
                currentDay = 1;
                currentMonth++;
            }
            else
                flag = false;

            if (flag)
                this.myFunPictureBox.Invalidate();
        }

        private void MonthLeft_Click(object sender, EventArgs e)
        {
            Boolean flag = true;
            if (currentMonth > 1)
                currentMonth--;
            else
                flag = false;

            if (flag)
                this.myFunPictureBox.Invalidate();
        }

        private void MonthRight_Click(object sender, EventArgs e)
        {
            Boolean flag = true;
            if (currentMonth < 12)
                currentMonth++;
            else
                flag = false;

            if (flag)
                this.myFunPictureBox.Invalidate();
        }

        private void AddDay_Click(object sender, EventArgs e)
        {
            Boolean flag = true;
            currentSpan++;
            if (currentSpan > 10)
            {
                currentSpan--;
                flag = false;
            }
            else
            {
                // 调整横坐标interval
            }

            if (flag)
                this.myFunPictureBox.Invalidate();
        }

        private void RemoveDay_Click(object sender, EventArgs e)
        {
            Boolean flag = true;
            currentSpan--;
            if (currentSpan < 1)
            {
                currentSpan++;
                flag = false;
            }
            else
            {
                // 调整横坐标interval
            }

            if (flag)
                this.myFunPictureBox.Invalidate();

        }

        private int getMaxDay(int year, int month)
        {
            switch (month)
            {
                case 1:
                case 3:
                case 5:
                case 7:
                case 8:
                case 10:
                case 12:
                    return 31;
                case 2:

                    if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
                        return 29;
                    else
                        return 28;
                default:
                    return 30;
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

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.isPanelDrag = true;
                this.panelPrePosition = new Point(e.X, e.Y);
            }
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && this.isPanelDrag)
            {
                this.panel1.Left = this.panel1.Left + e.X - this.panelPrePosition.X;
                this.panel1.Top = this.panel1.Top + e.Y - this.panelPrePosition.Y;
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.isPanelDrag = false;
            }
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
            Dictionary<string, Dictionary<string, string>> LogoParameter = new Dictionary<string, Dictionary<string, string>>();
            NORMData.Add(drawParameter);
            NORMData.Add(LogoParameter);
            int startRow = 0;
            while (!(NORMDt.Rows[startRow]["Item"].ToString().IndexOf("图例参数") > 0))
            {
                startRow++;
            }
            startRow++;
            while (!(NORMDt.Rows[startRow]["Item"].ToString().IndexOf("绘图参数") > 0))
            {
                if (LogoParameter.ContainsKey(NORMDt.Rows[startRow]["Flag"].ToString()))
                {
                    Dictionary<string, string> dictionary = LogoParameter[NORMDt.Rows[startRow]["Flag"].ToString()];
                    dictionary.Add("secondItem", NORMDt.Rows[startRow]["Item"].ToString());
                    dictionary.Add("secondHatch", NORMDt.Rows[startRow]["Hatch"].ToString());
                    dictionary.Add("secondARGB", NORMDt.Rows[startRow]["ARGB"].ToString());
                    dictionary.Add("secondMark", "1");
                }
                else
                {
                    Dictionary<string, string> dictionary = new Dictionary<string, string>();
                    dictionary.Add("firstItem", NORMDt.Rows[startRow]["Item"].ToString());
                    dictionary.Add("firstHatch", NORMDt.Rows[startRow]["Hatch"].ToString());
                    dictionary.Add("firstARGB", NORMDt.Rows[startRow]["ARGB"].ToString());
                    dictionary.Add("firstMark", "1");
                    LogoParameter.Add(NORMDt.Rows[startRow]["Flag"].ToString(), dictionary);
                }
                startRow++;
            }
            startRow++;
            for (; startRow < NORMDt.Rows.Count; startRow++)
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
