using DevComponents.DotNetBar;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace HUST_Grph
{
    public partial class 画图 : Form
    {
        //private progress myprogress;
        private MyFunPictureBox myFunPictureBox = null;              //保存图片框控件
        private List<Dictionary<string, string>> STYLlist;
        private List<Dictionary<string, Dictionary<string, string>>> NORMlist;
        private Dictionary<string, Dictionary<string, List<string>>> MAPDictionary_Line;
        private Dictionary<string, Dictionary<string, List<string>>> MAPDictionary_Day;
        private string title;
        private string defaultPath = @"D:\HUST_Grph\";
        private int[] coordinateList;
        //private DrawHelper drawHelper;
        private Int32 currentDay;
        private Int32 currentMonth;
        private Int32 currentYear;
        private Int32 currentSpan;
        private Dictionary<string, HatchStyle> hatchStyles = new Dictionary<string, HatchStyle>();
        //private HatchStyle[] hatchStyles = new HatchStyle[18];

        private Boolean isPanelDrag = false;
        //private Boolean isMouseMove = false;
        private Boolean isLogoOn = true;
        private Boolean isPanelOn = true;
        private Boolean isHorizontal = true;
        private Boolean isTooltipOn = true;
        private int tooltipDay = -1;
        private int tooltipHour = -1;
        private string tooltipText = "";
        private System.Windows.Forms.Timer tooltipTimer;
        private System.Windows.Forms.Timer hoverRefreshTimer;
        //private Boolean tooltipHover = false;

        //private Boolean isHoverPic = true;
        //private Boolean picLandScape = true;
        private Point panelPrePosition;
        private Point start;
        private int[] baseLoad = new int[366];

        private Boolean isDirectSave = false;

        private Point preLeftTop;
        private DateTime lastMouseMove;
        private String path = null;


        [DllImport("user32")]
        private static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, IntPtr lParam);
        private const int WM_SETREDRAW = 0xB;

        public 画图(DataSet ds)
        {
            InitializeComponent();
            //string path = "D:\\input_11_2022020000_70_0.xlsx";

            loadData LoadData = new loadData();
            
            STYLlist = loadData.STYLTableToData(ds.Tables["STYL"]);
            NORMlist = loadData.NORMTableToData(ds.Tables["NORM"], baseLoad);
            MAPDictionary_Line = loadData.MAPsTableToData_Line(ds.Tables["Maps"]);
            MAPDictionary_Line = (from d in MAPDictionary_Line orderby d.Key descending select d).ToDictionary(k => k.Key, v => v.Value);
            MAPDictionary_Day = loadData.MAPsTableToData_Day(ds.Tables["Maps"]);
            //DataColumn[] datacolumn = {
            //    ds.Tables["NORM"].Columns["基荷A"],
            //    ds.Tables["NORM"].Columns["基荷B"],
            //};
            int[] baseline = new int[366];


            title = STYLlist[2]["5"];
            path = STYLlist[2]["9"];
            dlgSavePic.FileName = title;

            tooltipTimer = new System.Windows.Forms.Timer();
            tooltipTimer.Interval = 500; // 设置定时器的间隔为0.5秒
            tooltipTimer.Tick += TooltipTimer_Tick;

            hoverRefreshTimer = new System.Windows.Forms.Timer();
            hoverRefreshTimer.Interval = 30000; // 设置定时器的间隔为30秒
            hoverRefreshTimer.Tick += hoverRefreshTimer_Tick;

            Dictionary<String, String> hatchDictionary = STYLlist[1];

            foreach (var pair in hatchDictionary)
            {
                HatchStyle hatchStyle;
                string key = pair.Key;
                string value = pair.Value;
                if (value == "无填充")
                {
                    continue;
                }
                hatchStyle = (HatchStyle)Enum.Parse(typeof(HatchStyle), value);
                hatchStyles.Add(key, hatchStyle);
            }


        }
        public 画图(DataSet ds, Boolean tf0, String filePath)
        {
            InitializeComponent();
            //this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);

            //string path = "D:\\input_11_2022020000_70_0.xlsx";
            loadData LoadData = new loadData();

            STYLlist = loadData.STYLTableToData(ds.Tables["STYL"]);
            NORMlist = loadData.NORMTableToData(ds.Tables["NORM"], baseLoad);
            MAPDictionary_Line = loadData.MAPsTableToData_Line(ds.Tables["Maps"]);
            MAPDictionary_Line = (from d in MAPDictionary_Line orderby d.Key descending select d).ToDictionary(k => k.Key, v => v.Value);
            MAPDictionary_Day = loadData.MAPsTableToData_Day(ds.Tables[2]);

            title = STYLlist[2]["5"];
            path = STYLlist[2]["9"];
            dlgSavePic.FileName = filePath;

            // 无需显示数据点
            //tooltipTimer = new System.Windows.Forms.Timer();
            //tooltipTimer.Interval = 500; // 设置定时器的间隔为0.5秒
            //tooltipTimer.Tick += TooltipTimer_Tick;

            //hoverRefreshTimer = new System.Windows.Forms.Timer();
            //hoverRefreshTimer.Interval = 30000; // 设置定时器的间隔为30秒
            //hoverRefreshTimer.Tick += hoverRefreshTimer_Tick;

            Dictionary<String, String> hatchDictionary = STYLlist[1];

            foreach (var pair in hatchDictionary)
            {
                HatchStyle hatchStyle;
                string key = pair.Key;
                string value = pair.Value;
                if (value == "无填充")
                {
                    continue;
                }
                hatchStyle = (HatchStyle)Enum.Parse(typeof(HatchStyle), value);
                hatchStyles.Add(key, hatchStyle);
            }

            // 设置图例显示
            this.isDirectSave = true;
            this.isLogoOn = tf0;
            this.newTab();
            this.SavePicMenuItem.PerformClick();
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

        public void newTab()
        {
            TabItem tp = this.tabControl1.CreateTab("图1");

            TabControlPanel tcp = new TabControlPanel();
            tcp.Visible = false;
            tcp.TabItem = tp;
            tcp.AutoScroll = true;
            tcp.Dock = DockStyle.Fill;

            tcp.Resize += new System.EventHandler(this.pictureBox_Resize);

            currentYear = int.Parse(STYLlist[2]["3"].Substring(0, 4));
            currentDay = int.Parse(STYLlist[2]["15"]);
            currentMonth = int.Parse(STYLlist[2]["14"]);
            currentSpan = int.Parse(STYLlist[2]["16"]);

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

            //初始修改为横向 cjy 7.27
            myFunPictureBox.Width = (int)(pageSettings.PaperSize.Height / 100.0 * 96);
            myFunPictureBox.Height = (int)(pageSettings.PaperSize.Width / 100.0 * 96);

            int xDraw = (int)(pageSettings.Margins.Top / 254.0 * 96);
            int yDraw = (int)(pageSettings.Margins.Left / 254.0 * 96);
            int widthDraw =
            (int)(pageSettings.PaperSize.Height / 100.0 * 96) -
                (int)((pageSettings.Margins.Top + pageSettings.Margins.Bottom) / 254.0 * 96);
            int heightDraw = (int)(pageSettings.PaperSize.Width / 100.0 * 96) -
                (int)((pageSettings.Margins.Left + pageSettings.Margins.Right) / 254.0 * 96);
            myFunPictureBox.drawArea = new Rectangle(xDraw, yDraw, widthDraw, heightDraw);

            myFunPictureBox.logoPos = new Rectangle(myFunPictureBox.drawArea.Left + (int)(myFunPictureBox.drawArea.Width * 0.5)
                , myFunPictureBox.drawArea.Bottom - (int)(myFunPictureBox.drawArea.Height * 0.4)
                , 300
                , (myFunPictureBox.LogoItems.Count + 1) / 2 * 50);

            myFunPictureBox.ContextMenuStrip = contextMenuStrip1;

            myFunPictureBox.MouseWheel += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseWheel);
            myFunPictureBox.MouseLeave += new System.EventHandler(this.pictureBox_MouseLeave);
            myFunPictureBox.MouseEnter += new System.EventHandler(this.pictureBox_MouseEnter);
            myFunPictureBox.MouseMove += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseMove);
            myFunPictureBox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseDown);
            myFunPictureBox.MouseUp += new System.Windows.Forms.MouseEventHandler(this.pictureBox_MouseUp);
            //myFunPictureBox.MouseHover += new System.EventHandler(this.pictureBox_Hover);

            myFunPictureBox.KeyDown += new System.Windows.Forms.KeyEventHandler(this.myFunPictureBox_KeyDown);
            myFunPictureBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.myFunPictureBox_KeyUp);

            myFunPictureBox.Paint += new System.Windows.Forms.PaintEventHandler(this.pictureBox_Paint);
            myFunPictureBox.Resize += new System.EventHandler(this.pictureBox_Resize);
            //myFunPictureBox.LostFocus += new System.EventHandler(this.);


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
            //this.toolTip2.Active = true;
        }

        private void pictureBox_MouseLeave(object sender, System.EventArgs e)
        {
            // this.Focus();
            this.toolTip2.Active = false;
            this.tooltipTimer.Stop();
        }

        //指示是否正在进行图例拖动
        private bool isLogoDrag = false;
        //指示是否正在进行图像推动
        private bool isPicDrag = false;

        private void pictureBox_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            MyPictureBox picture = sender as MyPictureBox;
            picture.previousPos = new Point(e.X, e.Y);

            if (e.Button == MouseButtons.Left)
            {
                //指示开始图例拖动
                if (this.isLogoOn && picture.logoPos.Contains(picture.previousPos))
                    this.isLogoDrag = true;
                //指示开始图像拖动
                else if (picture.DisplayRectangle.Contains(picture.previousPos))
                {
                    this.isPicDrag = true;
                }

            }
        }

        private void pictureBox_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {

            //指示结束图例/图像拖动
            if (e.Button == MouseButtons.Left)
            {
                this.isLogoDrag = false;
                this.isPicDrag = false;
            }
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
                panel.AutoScroll = true;
            }
            else
            {
                Point mousePoint = new Point(e.X, e.Y);
                Panel panel = (sender as PictureBox).Parent as Panel;
                mousePoint.Offset(this.Location.X, this.Location.Y);
                if (panel.RectangleToScreen(panel.DisplayRectangle).Contains(mousePoint))
                {
                    panel.AutoScrollPosition = new Point(panel.HorizontalScroll.Value, panel.VerticalScroll.Value - e.Delta);
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

                paintDays(g, picture, currentYear, currentMonth, currentDay, currentSpan);

                if (isLogoOn == true)
                {
                    DrawLogo(picture, g);
                }

                picture.drawed = true;


                preLeftTop = new Point(myFunPictureBox.Bounds.X, myFunPictureBox.Bounds.Y);
            }
            catch (Exception ex)
            {
                //ex.WriteLog();
                MessageBox.Show(ex.ToString());
            }

        }

        // 开始绘图的年、月、日、绘图天数、一年最大值 
        private void paintDays(Graphics g, MyFunPictureBox picture, int year, int startMonth, int startDay, int days)
        {

            // 左右边界的宽度
            int leftBorderWidth = (int)(picture.drawArea.Width * 0.1);
            int rightBorderWidth = (int)(picture.drawArea.Width * 0.01);
            int topBorderWidth = (int)(picture.drawArea.Height * 0.05);
            int bottomBorderWidth = (int)(picture.drawArea.Height * 0.01);

            // 整个画面的宽度和高度
            int wholeWidth = this.myFunPictureBox.drawArea.Width;
            int wholeHeight = this.myFunPictureBox.drawArea.Height;

            int width = wholeWidth - leftBorderWidth - rightBorderWidth - 10;

            //根据图片的绘制宽度调整字体的大小,并且保证字体的大小在一定范围内，例如7-14
            float fontSize = width / 70 - 2;
            if (fontSize > 12)
            {
                fontSize = 12;
            }
            else if (fontSize < 6)
            {
                fontSize = 6;
            }
            int hours = 24 * days;

            //由于整数存在缩进，需要对不能进行整除的部分进行处理，选择的方式是将多余的部分平均分配到每一天的前i个小时中
            int hourLeft = width % hours;
            int dayRemainder = hourLeft % days;
            int prehours = hourLeft / days;
            int hourWidth = width / hours;

            coordinateList = new int[hours + 1];

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



            // 从STYL表中获取最大纵坐标
            Dictionary<string, string> STYLinfo1 = STYLlist[2];
            int maxVal = int.Parse(STYLinfo1["12"]);
            int minVal = int.Parse(STYLinfo1["17"]);


            // 绘制坐标轴
            Pen pen = new Pen(Brushes.Black);
            pen.Width = 1.8F;
            // 绘制X轴
            start = new Point(leftBorderWidth, wholeHeight - bottomBorderWidth);
            Point end = new Point(wholeWidth - rightBorderWidth, wholeHeight - bottomBorderWidth);
            g.DrawLine(pen, start, end);
            Point[] arrow = new Point[2];
            arrow[0] = new Point(end.X - 5, end.Y + 3);
            arrow[1] = new Point(end.X - 5, end.Y - 3);
            g.DrawLine(pen, arrow[0], end);
            g.DrawLine(pen, arrow[1], end);

            // 绘制Y轴
            //y轴的顶点（最上方点）是end2
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
            Point[] baseLoadPoints = new Point[24 * days * 2];

            foreach (string key in this.MAPDictionary_Line.Keys)
            {
                //判断在绘图范围内是否存在该曲线，并根据该标志确定是否进行绘图
                Boolean isExist = false;
                string flag = key;
                Dictionary<string, List<string>> dic1 = MAPDictionary_Line[key];
                ArrayList array = new ArrayList();
                for (int i = 0; i < days; i++)
                {
                    string str = Convert.ToString(i + startDays);
                    if (dic1.ContainsKey(str))
                    {
                        isExist = true;
                        List<string> list = dic1[str];
                        foreach (string val in list)
                        {
                            array.Add(int.Parse(val));
                        }
                    }
                    else
                    {
                        for (int j = 0; j < 24; j++)
                        {
                            array.Add(0);
                        }
                    }
                }
                if (isExist)
                {
                    int[] data = (int[])array.ToArray(typeof(int));
                    // data-数据 end-绘图的左上角坐标 width-绘图的宽度 height-绘图的高度 maxVal-一年的最大值，用于固定Y轴
                    //Console.WriteLine(days + ":" + (wholeWidth - leftBorderWidth - rightBorderWidth - 5));

                    int height = start.Y - end2.Y - 5;


                    Point[] points = dataToPoint(data, start, coordinateList, height, maxVal);
                    // 笔刷中颜色、填充样式应该解析文件得到
                    Dictionary<string, Dictionary<string, string>> NORMdic = NORMlist[0];
                    Dictionary<string, string> NORMinfo = NORMdic[key];
                    string hatchID = NORMinfo["Hatch"];
                    string color = NORMinfo["ARGB"];
                    //Dictionary<string, string> STYLinfo = STYLlist[0];
                    //string color2 = STYLinfo[color];
                    //string[] ARGBS = color2.Split(' ');
                    if (key.Equals("9950"))
                    {
                        Array.Copy(points, originalLoadPoints, originalLoadPoints.Length); 
                    }
                    else
                    {
                        //HatchBrush htBrush = new HatchBrush(hatchStyles[hatchID], Color.Black, Color.FromArgb(int.Parse(ARGBS[0]), int.Parse(ARGBS[1]), int.Parse(ARGBS[2]), int.Parse(ARGBS[3])));
                        Brush hBrush = infoToBrush(color, hatchID);
                        // 画图
                        g.FillPolygon(hBrush, points);
                        g.DrawPolygon(Pens.Black, points);
                        hBrush.Dispose();
                    }
                }
            }
            startDays = dateToDay(year, startMonth, startDay);
            int[] baseLoaddata = new int[24 * days];
            for(int i = 0; i < days; i++)
            {
                for(int j = 0; j < 24; j++) {
                    int pointIndex = i * 24 + j;
                    baseLoaddata[pointIndex] = baseLoad[i + startDays];
                }
            }
            int _height = start.Y - end2.Y - 5; ;
            Array.Copy(dataToPoint(baseLoaddata, start, coordinateList, _height, maxVal), baseLoadPoints, baseLoadPoints.Length);
            //baseLoadPoints = dataToPoint(baseLoaddata, end2, coordinateList, _height, maxVal);
            string originalColor = STYLlist[0][NORMlist[0]["9950"]["ARGB"]];
            string[] orginalARGBS = originalColor.Split(' ');
            //string = STYLlist[0][];
            Pen dashPen = new Pen(Color.FromArgb(int.Parse(orginalARGBS[0]), int.Parse(orginalARGBS[1]), int.Parse(orginalARGBS[2]), int.Parse(orginalARGBS[3])), 1.0f);
            Pen LinePen = new Pen(Color.FromArgb(int.Parse(orginalARGBS[0]), int.Parse(orginalARGBS[1]), int.Parse(orginalARGBS[2]), int.Parse(orginalARGBS[3])), 1.0f);
            dashPen.DashStyle = DashStyle.Dash;
            //Pen blackPen = new Pen(Color.Black, (float)1.5);
            //blackPen.DashPattern = new float[] { 2, 1 };
            g.DrawLines(dashPen, originalLoadPoints);
            g.DrawLines(LinePen, baseLoadPoints);

            dashPen.Dispose();
            LinePen.Dispose();


            Point top;
            Pen dotted = new Pen(Color.Black, 1);
            dotted.DashStyle = System.Drawing.Drawing2D.DashStyle.Custom;
            dotted.DashPattern = new float[] { 1, 1 };

            // 绘制完图片部分，再去画坐标轴的间隔
            // 根据天数为1天还是多天，决定X轴分为24小时还是若干天
            if (days > 1)
            {
                //interval = (wholeWidth - leftBorderWidth - rightBorderWidth - 5) / (days*24) * 24;
                Point[] downPoints = new Point[days + 1];
                Point[] upPoints = new Point[days];

                String[] dates = getDateString(year, startMonth, startDay, days);

                int span = days / 10 + 1;

                for (int i = 0; i < days; i++)
                {
                    downPoints[i] = new Point(start.X + coordinateList[i * 24], start.Y);
                    upPoints[i] = new Point(start.X + coordinateList[i * 24], start.Y - 5);
                    top = new Point(start.X + coordinateList[i * 24], end2.Y + 8);

                    if (i % span == 0)
                    {
                        //绘制坐标轴刻度
                        g.DrawLine(Pens.Black, downPoints[i], upPoints[i]);
                        //绘制参考虚线
                        g.DrawLine(dotted, downPoints[i], top);
                        //坐标轴刻度大小
                        g.DrawString(dates[i], new Font("黑体", fontSize - 1), Brushes.Black, new Point(downPoints[i].X - 10, downPoints[i].Y + 5));
                    }

                }
                downPoints[days] = new Point(start.X + coordinateList[days * 24], start.Y);
                top = new Point(start.X + coordinateList[days * 24], end2.Y + 8);
                g.DrawLine(dotted, downPoints[days], top);

                //在图中打印横坐标轴的单位
                g.DrawString("T", new Font("黑体", fontSize + 1), Brushes.Black, downPoints[days].X + 15, downPoints[days].Y);
                //打印标题
                Rectangle rect;
                if (days % 2 == 0)
                {
                    rect = new Rectangle(downPoints[days / 2].X - 360, downPoints[days].Y + (int)(fontSize * 3.5), 720, 240);
                }
                else
                {
                    rect = new Rectangle((downPoints[days / 2].X + downPoints[days / 2 + 1].X) / 2 - 360, downPoints[days].Y + (int)(fontSize * 3.5), 720, 240);
                }

                var stringFormat = new StringFormat();
                stringFormat.Alignment = StringAlignment.Center;
                g.DrawString(title, new Font("黑体", fontSize + 2), Brushes.Black, rect, stringFormat);

                stringFormat.Dispose();

            }
            else
            {
                String[] dates = getDateString(year, startMonth, startDay, days);
                Font font = new Font("黑体", fontSize - 1);
                g.DrawString(dates[0], font, Brushes.Black, new Point(start.X - 28, start.Y + 5));
                //interval = (end.X - start.X - 5) / 24;
                Point[] downPoints = new Point[24];
                Point[] upPoints = new Point[24];
                for (int i = 0; i < 24; i++)
                {

                    downPoints[i] = new Point(start.X + coordinateList[i + 1], start.Y);
                    upPoints[i] = new Point(start.X + coordinateList[i + 1], start.Y - 5);
                    top = new Point(start.X + coordinateList[i + 1], end2.Y + 8);

                    if (i % 4 == 3)
                    {
                        g.DrawLine(dotted, downPoints[i], top);
                        g.DrawLine(Pens.Black, downPoints[i], upPoints[i]);
                        g.DrawString((i + 1) + "时", font , Brushes.Black, new Point(downPoints[i].X - 10, downPoints[i].Y + 5));
                    }
                }

                //在图中打印横坐标轴的单位
                g.DrawString("T", new Font("黑体", fontSize + 1), Brushes.Black, downPoints[23].X + 20, downPoints[days].Y - 2);
                //打印标题
                Rectangle rect = new Rectangle(downPoints[11].X - 360, downPoints[days].Y + (int)(fontSize * 3.5), 720, 80);
                var stringFormat = new StringFormat();
                stringFormat.Alignment = StringAlignment.Center;
                g.DrawString(title, new Font("黑体", fontSize + 2), Brushes.Black, rect, stringFormat);

                font.Dispose();
                stringFormat.Dispose();
            }

            // 绘制Y轴的间隔
            double intervalY = ((double)start.Y - end2.Y - 5) / 10;
            //wholeHeight - topBorderWidth - bottomBorderWidth - 5

            //从数据表中读取量纲
            float dimension = float.Parse(STYLinfo1["11"]);


            for (int i = 0; i < 10; i++)
            {
                StringFormat format = new StringFormat();
                format.Alignment = StringAlignment.Far; //靠右对齐
                Rectangle space = new Rectangle(start.X - 70, (start.Y - (int)(intervalY * (i + 1))), 70, 20);
                Point leftPoint = new Point(start.X, start.Y - (int)(intervalY * (i + 1)));
                Point rightPoint = new Point(start.X + 5, start.Y - (int)(intervalY * (i + 1)));
                top = new Point(end.X - 10, start.Y - (int)(intervalY * (i + 1)));
                g.DrawLine(Pens.Black, leftPoint, rightPoint);
                g.DrawLine(dotted, leftPoint, top);
                Font font = new Font("黑体", fontSize - 1);
                //根据量纲打印纵坐标轴的刻度
                g.DrawString((int)(maxVal * dimension / 10 * (i + 1)) + "", font, Brushes.Black, space, format);
                font.Dispose();
                format.Dispose();

            }

            //打印单位在图片上，打印的内容在STYL表中记录
            Rectangle unitSpace = new Rectangle(start.X - 35, end2.Y - 25, 80, 20);
            StringFormat unitFormat = new StringFormat();
            g.DrawString("P/" + STYLlist[2]["10"], new Font("黑体", fontSize + 1), Brushes.Black, unitSpace, unitFormat);

            dotted.Dispose();


        }

        // 数据转换为点坐标
        private Point[] dataToPoint(int[] data, Point start, int[] coordinateList, int height, int maxVal)
        {
            //Console.WriteLine(width);
            Point[] points = new Point[data.Length * 2 + 2];
            int days = data.Length / 24;

            int X = start.X;
            int Y = start.Y;
            double oneHeight = (float)height / maxVal;//每1单位对应的纵坐标长度
            int pointIndex = 0;
            for (int i = 0; i < days; i++)
            {
                for (int j = 0; j < 24; j++)//按照每天24小时进行作图
                {
                    int hours = i * 24 + j;//当前坐标点所表示的总小时位置
                    points[pointIndex] = new Point(X + coordinateList[hours], (int)(Y - data[hours] * (double)height / maxVal));
                    pointIndex++;
                    points[pointIndex] = new Point(X + coordinateList[hours + 1], (int)(Y - data[hours] * (double)height / maxVal ));
                    pointIndex++;
                }

            }

            points[points.Length - 2] = new Point(points[points.Length - 3].X, Y );
            points[points.Length - 1] = new Point(points[0].X, Y);
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
                if (month > 12 || month < 1)
                {
                    date[i] = "";
                }
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
            MyPictureBox picture = sender as MyPictureBox;
            // 处理图例拖动
            if (isLogoDrag == true)
            {
                //关闭ToolTip计时器
                tooltipTimer.Stop();
                toolTip2.Active = false;

                int width = picture.Bounds.Width - 100;
                int height = picture.Bounds.Height - 100;
                int PX = e.X < 0 ? 0 : e.X;
                PX = PX > width ? width : PX;
                int PY = e.Y < 0 ? 0 : e.Y;
                PY = PY > height ? height : PY;

                Point mousePoint = new Point(PX, PY);

                Point preZero = new Point(picture.logoPos.Left, picture.logoPos.Top);
                picture.logoPos = new Rectangle(
                    picture.logoPos.Left + PX - picture.previousPos.X,
                    picture.logoPos.Top + PY - picture.previousPos.Y,
                    picture.logoPos.Width,
                    picture.logoPos.Height);
                Point newZero = new Point(picture.logoPos.Left, picture.logoPos.Top);

                int redrawX = newZero.X > preZero.X ? preZero.X : newZero.X;
                int redrawY = newZero.Y > preZero.Y ? preZero.Y : newZero.Y;

                picture.Invalidate(new Rectangle(redrawX, redrawY, picture.Width, picture.Height));
                picture.previousPos = mousePoint;

                return;
            }
            // 处理图像拖动
            else if (isPicDrag)
            {
                //关闭ToolTip计时器
                tooltipTimer.Stop();
                toolTip2.Active = false;

                DateTime now = DateTime.Now;
                if ((now - lastMouseMove) < TimeSpan.FromMilliseconds(10))
                {
                    return; // 跳过处理
                }

                Panel panel = picture.Parent as Panel;

                int limitX = picture.previousPos.X - preLeftTop.X - picture.Width + 100;
                int limitY = picture.previousPos.Y - preLeftTop.Y - picture.Height + 100;
                int PX = e.X < limitX ? limitX : e.X;
                int PY = e.Y < limitY ? limitY : e.Y;

                int rangeX = picture.Width - panel.ClientRectangle.Width;
                int rangeY = picture.Height - panel.ClientRectangle.Height;

                int postX = preLeftTop.X + PX - picture.previousPos.X;
                int postY = preLeftTop.Y + PY - picture.previousPos.Y;

                if ((PX + rangeX) * (postX + rangeX) <= 0 || (PY + rangeY) * (postY + rangeY) <= 0)
                {
                    //禁止重绘
                    //SendMessage(panel.Handle, WM_SETREDRAW, 0, IntPtr.Zero);
                    panel.SuspendLayout();
                }
                else
                {
                    //SendMessage(panel.Handle, WM_SETREDRAW, 1, IntPtr.Zero);
                    panel.ResumeLayout();
                }

                picture.Bounds = new Rectangle(
                    preLeftTop.X + PX - picture.previousPos.X,
                    preLeftTop.Y + PY - picture.previousPos.Y,
                    picture.Bounds.Width,
                    picture.Bounds.Height);


                preLeftTop = new Point(picture.Bounds.X, picture.Bounds.Y);
                //picture.previousPos = new Point(PX, PY);
                lastMouseMove = now; // 更新上次处理事件的时间

                picture.Invalidate();
            }
            // 鼠标悬停于图例内不按下
            else if (picture.logoPos.Contains(new Point(e.X, e.Y)))
            {
                tooltipTimer.Stop();
            }
            // 处理数据点显示
            else if (this.isTooltipOn)
            {
                Point cursorPosition = e.Location;

                int day = 0;
                int hour = 0;

                if (coordinateList == null)
                {
                    return;
                }
                else
                {
                    int toolPointX = cursorPosition.X - start.X;
                    int maxRangeX = coordinateList[coordinateList.Length - 1];
                    if (0 > toolPointX || toolPointX > maxRangeX)
                    {
                        tooltipDay = 0;
                        tooltipHour = 0;
                        tooltipTimer.Stop();
                        toolTip2.Active = false;
                    }
                    else
                    {
                        //toolTip2.Active = true;
                        for (int i = 1; i < coordinateList.Length; i++)
                        {
                            if (toolPointX < coordinateList[i])
                            {
                                day = (i - 1) / 24;
                                hour = (i - 1) % 24;

                                if (day == tooltipDay && hour == tooltipHour)
                                {
                                    //数据点不变，不处理

                                    return;
                                }
                                else
                                {

                                    toolTip2.Active = false;

                                    tooltipDay = day;
                                    tooltipHour = hour;

                                    //数据点变化，重置显示计时器，关闭刷新计时器
                                    tooltipTimer.Stop();
                                    tooltipTimer.Start();
                                    hoverRefreshTimer.Stop();

                                }

                                break;
                            }
                        }
                    }
                }
            }
        }
        private void TooltipTimer_Tick(object sender, EventArgs e)
        {
            tooltipTimer.Stop();
            //计时10秒后刷新，避免ToolTip关闭
            hoverRefreshTimer.Start();

            string[] dates = getDateString(currentDay, currentMonth, currentDay, currentSpan);
            int day = dateToDay(currentYear, currentMonth, currentDay) + tooltipDay;
            int hour = tooltipHour;
            string text = "";
            //吴老师需要将出力合计在一起
            //公式为
            //出力 = 基荷 + 腰荷 + 峰荷
            //因此先统计包含哪些需要显示在tooltip中的数据，并获取名称保存在total的key中
            //再依次将数据存储在total的value中
            Dictionary<String, int> total = new Dictionary<string, int>();
            Dictionary<String, List<String>> data;

            if (!MAPDictionary_Day.ContainsKey(day.ToString()))
            {

                return;
            }
            data = MAPDictionary_Day[day.ToString()];
            data = (from d in data orderby d.Key descending select d).ToDictionary(k => k.Key, v => v.Value);

            string prename = "";
            foreach (var pair in data)
            {
                string key = pair.Key;
                int firstNum = int.Parse(key.Substring(0, 1));
                int secondNum = int.Parse(key.Substring(1, 1));
                int thirdNum = int.Parse(key.Substring(2, 1));
                int fourthNum = int.Parse(key.Substring(3, 1));
                string name = "";
                if (firstNum <= 3)
                {
                    if (fourthNum == 0)
                    {
                        switch (secondNum)
                        {
                            case 0:
                                name = "核电出力";
                                break;
                            case 1:
                                name = "火电出力";
                                break;
                            case 2:
                                name = "水电出力";
                                break;
                            case 3:
                                name = "储能出力";
                                break;
                            default:
                                Console.WriteLine("flag超出原定绘图参数");
                                break;

                        }
                    }
                    else
                    {
                        if (fourthNum > 5)
                        {
                            Console.WriteLine("电站超出5个");
                        }
                        else
                        {
                            name = STYLlist[3][fourthNum.ToString()].Split('*')[1];
                            switch (secondNum)
                            {
                                case 0:
                                //吴老师要求将“核力出力”等改成“发”，缩短宽度
                                //name += "核电出力";
                                //break;
                                case 1:
                                //name += "水电出力";
                                //break;
                                case 2:
                                //name += "火电出力";
                                //break;
                                case 3:
                                    //name += "储能出力";
                                    //break;
                                    name += "发";
                                    break;
                                default:
                                    Console.WriteLine("flag超出原定绘图参数");
                                    break;

                            }
                        }
                    }
                }
                else
                {
                    if (fourthNum == 0)
                    {
                        name = NORMlist[0][key]["Item"];
                    }
                    else
                    {
                        if (fourthNum > 5)
                        {
                            Console.WriteLine("电站超出5个");
                        }
                        else
                        {
                            name = STYLlist[3][fourthNum.ToString()].Split('*')[1];
                            //name += NORMlist[0][key.Remove(3, 1) + "0"]["Item"];
                            Dictionary<int, String> namePairs = new Dictionary<int, string>{
                                { 4, "储" },
                                { 5, "发" },
                                { 7, "弃" }
                            };
                            name += namePairs[firstNum];
                        }
                    }

                }
                //需要用flag高一级的减去flag第一级的电量获得该flag对应的电量
                if (prename != "原始负荷" && prename != "")
                {
                    total[prename] = total[prename] - int.Parse(pair.Value[hour]);
                }
                if (total.ContainsKey(name))
                {
                    total[name] += int.Parse(pair.Value[hour]);
                }
                else
                {
                    total.Add(name, int.Parse(pair.Value[hour]));
                }
                prename = name;
            }

            foreach (var pair in total)

            {
                text += "   ";
                text += pair.Key;
                text += "： ";
                text += pair.Value;
                text += "   \n";
            }
            text += " （上述电源出力不含被\n    储能负荷覆盖的部分）";

            this.tooltipText = text;
            this.toolTip2.ToolTipTitle = "   " + dates[tooltipDay] + "  " + (hour + 1) + "时/MW";
            this.toolTip2.SetToolTip(this.myFunPictureBox, this.tooltipText);
            this.toolTip2.Active = true;            

        }

        private void hoverRefreshTimer_Tick(object sender, EventArgs e)
        {
            //并不关闭，反复刷新
            //hoverRefreshTimer.Stop();
            this.toolTip2.SetToolTip(this.myFunPictureBox, this.tooltipText);
        }

        private string WrapLogoString(string originalStr)
        {
            string result = "";
            for (int i = 0; i < originalStr.Length; i++)
            {
                if (originalStr[i] == '\\')
                {
                    result += "\n";
                    continue;
                }
                result += originalStr[i];
            }
            return result;
        }
        //private void SortLogo(MyPictureBox picture)
        //{
        //    for (int i = 0; i < picture.LogoItems.Count; i++)
        //    {
        //        int minPriority = 10000;
        //        int index = -1;
        //        for (int j = i; j < picture.LogoItems.Count; j++)
        //            if (picture.LogoItems[j].priority < minPriority)
        //            {
        //                minPriority = picture.LogoItems[j].priority;
        //                index = j;
        //            }
        //        LogoItem item = new LogoItem();
        //        item.brush = picture.LogoItems[index].brush;
        //        item.description = picture.LogoItems[index].description;
        //        item.priority = picture.LogoItems[index].priority;
        //        picture.LogoItems.RemoveAt(index);
        //        picture.LogoItems.Insert(i, item); ;
        //    }
        //}
        private void DrawLogo(MyPictureBox picture, Graphics g)
        {
            LogoItem newItem = new LogoItem();
            newItem.brush = new SolidBrush(Color.SkyBlue);
            newItem.description = "原始负荷";
            picture.LogoItems.Add(newItem);
            //    newItem.priority = 0;

            Dictionary<String, Dictionary<String, String>> LogoParameter = NORMlist[1];

            for (int i = 0; i < LogoParameter.Count; i++)
            {
                KeyValuePair<String, Dictionary<String, String>> kv = LogoParameter.ElementAt(i);
                Dictionary<String, String> Info = kv.Value;
                int firstMark = int.Parse(Info["firstMark"]);
                int secondMark = int.Parse(Info["secondMark"]);

                string firstItem = Info["firstItem"];
                string secondItem = Info["secondItem"];

                if ((firstMark == 0 && secondMark == 0) || firstItem.Equals("原始负荷"))
                {
                    continue;
                }
                string firstHatchID = Info["firstHatch"];
                string secondHatchID = Info["secondHatch"];
                string firstColor = Info["firstARGB"];
                string secondColor = Info["secondARGB"];

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

                    newItem.description = firstItem + "\\" + secondItem;

                }
                picture.LogoItems.Add(newItem);
            }


            SolidBrush backBrush = new SolidBrush(Color.White);
            Font drawFont = new Font("宋体", picture.smallFontSize);

            SolidBrush drawBrush = new SolidBrush(Color.Black);
            Pen framePen = new Pen(Color.Black, 1.0f);
            string originalColor = STYLlist[0][NORMlist[0]["9950"]["ARGB"]]; ;
            string[] orginalARGBS = originalColor.Split(' ');
            //string = STYLlist[0][];
            Pen dashPen = new Pen(Color.FromArgb(int.Parse(orginalARGBS[0]), int.Parse(orginalARGBS[1]), int.Parse(orginalARGBS[2]), int.Parse(orginalARGBS[3])), 1.0f);
            dashPen.DashStyle = DashStyle.Dash;
            //Pen dashPen = new Pen(Color.Black, 1.0f);
            //dashPen.DashStyle = DashStyle.Dash;

            int vacant = 5;
            Font titleFont = new Font("宋体", picture.largeFontSize, FontStyle.Bold);


            //picture.logoPos = new Rectangle(picture.logoPos.Left
            //    , picture.logoPos.Top
            //    , picture.logoWidth
            //    , (picture.LogoItems.Count + 1) / 2 * (drawFont.Height * 2 + vacant) +
            //    vacant + titleFont.Height + vacant + 3);

            picture.logoPos = new Rectangle(picture.logoPos.Left
                , picture.logoPos.Top
                , picture.logoWidth
                , (picture.LogoItems.Count + 2) / 3 * (drawFont.Height * 2 + vacant) +
                vacant + 3);

            g.FillRectangle(backBrush, picture.logoPos);
            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;

            int itemWidth = picture.logoPos.Width / 3;

            //g.DrawString("图例", titleFont, drawBrush, picture.logoPos.Left + picture.logoPos.Width / 2, picture.logoPos.Top + vacant, stringFormat);
            Pen pen = new Pen(Color.Black);
            //g.DrawLine(pen, picture.logoPos.Left, picture.logoPos.Top + vacant + titleFont.Height + 3,
            //        picture.logoPos.Right, picture.logoPos.Top + vacant + titleFont.Height + 3);
            Point startPoint = new Point(picture.logoPos.Left, picture.logoPos.Top + 3);

            for (int i = 0; i < picture.LogoItems.Count; i++)
            {
                Point point = new Point(startPoint.X + 5 + (i % 3) * itemWidth,
                    startPoint.Y + (i / 3) * (drawFont.Height * 2 + vacant) + vacant);
                if (i == 0)
                {
                    g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y + drawFont.Height, 30, drawFont.Height);
                    g.DrawLine(framePen, point.X, point.Y + drawFont.Height, point.X + 30, point.Y + drawFont.Height);

                    PointF[] points = new PointF[]
                    {
                                new PointF(point.X, point.Y + (float)drawFont.Height*3.0f/2),
                                new PointF(point.X + 15, point.Y + (float)drawFont.Height*3.0f/2),
                                new PointF(point.X + 15, point.Y+(float)drawFont.Height/2.0f),
                                new PointF(point.X + 30, point.Y+(float)drawFont.Height/2.0f)
                        };
                    g.DrawLines(dashPen, points);
                    g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 35, point.Y);
                }
                else
                {
                    g.FillRectangle(picture.LogoItems[i].brush, point.X, point.Y, 30, drawFont.Height);
                    g.FillRectangle(picture.LogoItems[i].secondBrush, point.X, point.Y + drawFont.Height, 30, drawFont.Height);
                    g.DrawString(WrapLogoString(picture.LogoItems[i].description), drawFont, drawBrush, point.X + 35, point.Y);
                }
                g.DrawRectangle(framePen, point.X, point.Y, 30, drawFont.Height * 2);
            }

            stringFormat.Dispose();
            pen.Dispose();
            framePen.Dispose();
            drawBrush.Dispose();
            titleFont.Dispose();
            drawFont.Dispose();
        }

        //通过NORM表格中的color和hatch的编码，在Style表格中获取对应的填充风格和颜色类型，并生成对应的笔刷
        private Brush infoToBrush(string colorID, string hatchID)
        {
            Dictionary<string, string> STYLinfo = STYLlist[0];


            //当hatchID=39，也就是超出了hatchStyles的列表范围时，代表的是无填充风格，因此生成只包含颜色的SolidBrush笔刷，否则生成带有填充风格的HatchBrush笔刷
            if (hatchStyles.ContainsKey(hatchID))
            {
                string color = STYLinfo[colorID];
                string[] ARGBS = color.Split(' ');
                HatchBrush hBrush = new HatchBrush(hatchStyles[hatchID], Color.Black, Color.FromArgb(int.Parse(ARGBS[0]), int.Parse(ARGBS[1]), int.Parse(ARGBS[2]), int.Parse(ARGBS[3])));
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
            if (dateToDay(currentYear, currentMonth, currentDay) + currentSpan < getMaxDay(currentYear, 2) + 337)
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
            Boolean flag = false;
            if (dateToDay(currentYear, currentMonth + 1, currentDay) + currentSpan < getMaxDay(currentYear, 2) + 337)
            {
                flag = true;
                currentMonth++;
                //if (currentMonth < 12)
                //    currentMonth++;
                //else
                //    flag = false;
            }
            else
            {
                currentMonth = 12;
                currentDay = 31 - currentSpan + 1;
                flag = true;
  
            }
            if (flag)
                this.myFunPictureBox.Invalidate();

        }

        private void AddDay_Click(object sender, EventArgs e)
        {
            
            if (dateToDay(currentYear, currentMonth, currentDay) + currentSpan< getMaxDay(currentYear , 2) + 337)
            {
                Boolean flag = true;

                currentSpan++;

                if (flag)
                    this.myFunPictureBox.Invalidate();
            }
            
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
        private void InZoomPic(MyPictureBox picture, Panel panel)
        {
            if (picture.Width * 1.1 <= 2400 && picture.Height * 1.1 <= 3400)
            {
                picture.Size = new Size((int)(picture.Width * 1.1), (int)(picture.Height * 1.1));
                panel.AutoScroll = true;
                picture.smallFontSize = (float)(picture.smallFontSize * 1.1);
                picture.largeFontSize = (float)(picture.largeFontSize * 1.1);
                picture.logoWidth = (int)(picture.logoWidth * 1.1);
                picture.logoPos = new Rectangle((int)(picture.logoPos.Left * 1.1)
                    , (int)(picture.logoPos.Top * 1.1)
                    , picture.logoPos.Width
                    , picture.logoPos.Height);
                picture.drawArea = new Rectangle((int)(picture.drawArea.Left * 1.1),
                      (int)(picture.drawArea.Top * 1.1),
                      (int)(picture.drawArea.Width * 1.1),
                      (int)(picture.drawArea.Height * 1.1));
                picture.Bounds = new Rectangle(
                    preLeftTop.X, preLeftTop.Y,
                    picture.Width, picture.Height);
                picture.Invalidate();
            }
        }

        private void OutZoomPic(MyPictureBox picture, Panel panel)
        {
            double newWidth = picture.Width * 0.9;
            double newHeight = picture.Height * 0.9;
            if (newWidth > 300 && newHeight > 500 && preLeftTop.X + newWidth > 100 && preLeftTop.Y + newHeight > 100)
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
                picture.Bounds = new Rectangle(
                    preLeftTop.X, preLeftTop.Y,
                    picture.Width, picture.Height);
                picture.Invalidate();
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            //this.isMouseMove = false;
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
                //this.isMouseMove = true;
            }

        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.isPanelDrag = false;
            }
        }

        //private void button_MouseUp(object sender, MouseEventArgs e)
        //{
        //    if (!isMouseMove)
        //    {

        //    }
        //}


        private void SavePicMenuItem_Click(object sender, EventArgs e)
        {
            string parentDirPath = "";
            string filename = dlgSavePic.FileName;  
            String picPath = !isDirectSave ? defaultPath + filename + @".jpg" : filename;

            try
            {
                FileInfo fileInfo = new FileInfo(picPath);
                parentDirPath = fileInfo.DirectoryName;
                if (Directory.Exists(parentDirPath) == false) // 如果父亲文件夹不存在则创建
                {
                    Directory.CreateDirectory(parentDirPath);
                }
            }
            catch
            {
                Console.WriteLine("创建父目录失败");
            }

            if (isDirectSave)
            {
                MyFunPictureBox picture = this.myFunPictureBox;
                if (SavePicture(picture, dlgSavePic.FileName) == true)
                {
                    Console.WriteLine("SheetName is null");
                    // 直接保存时不弹窗
                    //MessageBox.Show("保存图片成功！", "提示",
                    //    MessageBoxButtons.OK,
                    //    MessageBoxIcon.Information,
                    //    MessageBoxDefaultButton.Button1,
                    //    MessageBoxOptions.DefaultDesktopOnly);
                }
            }
            else
            {
                dlgSavePic.InitialDirectory = parentDirPath;
                if (dlgSavePic.ShowDialog() == DialogResult.OK)
                {

                    ////开启进度条
                    //Thread thdSub = new Thread(new ThreadStart(this.progressB));
                    //thdSub.Start();
                    //Thread.Sleep(100);
                    MyFunPictureBox picture = ((sender as ToolStripMenuItem).Owner as ContextMenuStrip).SourceControl as MyFunPictureBox;

                    ////关闭进度条
                    //this.myprogress.isOver = true;

                    SavePicture(picture, dlgSavePic.FileName);
                    MessageBox.Show("保存图片成功！", "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);

                }
                dlgSavePic.FileName = title;
            }

        }

        private Boolean SavePicture(MyFunPictureBox picture, string filename)
        {

            MyFunPictureBox picBox = tabControl1.SelectedPanel.Controls[0] as MyFunPictureBox;
            //MyFunPictureBox imagePic = CreatePictureFromSource(picBox);
            MyFunPictureBox imagePic = picBox;

            Bitmap memImage = new Bitmap(imagePic.Width, imagePic.Height);

            Graphics g = Graphics.FromImage(memImage);
            imagePic.LogoItems.Clear();

            g.Clear(Color.White);

            paintDays(g, picture, currentYear, currentMonth, currentDay, currentSpan);

            //myDrawHelper.drawAxes(picture, g);

            if (isLogoOn == true)
            {
                DrawLogo(picture, g);
            }

            //filename由调用参数给出
            //String picPath = this.defaultPath + filename + @".jpg";
            String picPath = filename == title ? defaultPath + filename + @".jpg" : filename;

            //DialogResult result = DialogResult.OK;
            // 直接保存时不询问
            //if (isDirectSave == false)
            //{
            //    if (File.Exists(picPath))
            //    {
            //        result = MessageBox.Show("已有同名文件，是否覆盖？", "保存图片", MessageBoxButtons.OKCancel);
            //    }

            //    if (result == DialogResult.Cancel)
            //    {
            //        return false;
            //    }
            //}

            try
            {
                memImage.Save(picPath);
            }

            catch (Exception ex)
            {
                Console.WriteLine($"保存失败：{ex.Message}");
                throw;
            }

            g.Dispose();
            memImage.Dispose();

            return true;
        }

        private void 隐藏显示图例ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isLogoOn = !isLogoOn;
            隐藏显示图例ToolStripMenuItem.Text = isLogoOn ? "隐藏图例" : "显示图例";
            this.myFunPictureBox.Invalidate();
        }

        private void 隐藏显示控制板ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isPanelOn = !isPanelOn;
            隐藏显示控制板ToolStripMenuItem.Text = isPanelOn ? "隐藏控制板" : "显示控制板";
            this.panel1.Visible = isPanelOn;
        }

        private void 切换宽高比ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int tempWidth = myFunPictureBox.Width;
            myFunPictureBox.Width = myFunPictureBox.Height;
            myFunPictureBox.Height = tempWidth;

            PageSettings pageSettings = myFunPictureBox.pageSettings;

            // 根据新的宽高重新计算drawArea的边界坐标
            int xDraw = (int)(pageSettings.Margins.Left / 254.0 * 96);
            int yDraw = (int)(pageSettings.Margins.Top / 254.0 * 96);
            int widthDraw = myFunPictureBox.Width - (int)((pageSettings.Margins.Left + pageSettings.Margins.Right) / 254.0 * 96);
            int heightDraw = myFunPictureBox.Height - (int)((pageSettings.Margins.Top + pageSettings.Margins.Bottom) / 254.0 * 96);
            myFunPictureBox.drawArea = new Rectangle(xDraw, yDraw, widthDraw, heightDraw);

            isHorizontal = !isHorizontal;
            切换宽高比ToolStripMenuItem.Text = isHorizontal ? "纵向显示" : "横向显示";
            this.myFunPictureBox.Invalidate();
        }

        private void 隐藏显示数据点ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            isTooltipOn = !isTooltipOn;
            隐藏显示数据点ToolStripMenuItem.Text = isTooltipOn ? "隐藏数值" : "显示数值";
            this.toolTip2.Active = isTooltipOn;
            this.tooltipTimer.Stop();
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
                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                }
                else
                {
                    Console.WriteLine("SheetName is null");
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    if (sheetName.Equals("STYL"))
                    {
                        cellCount = 3;
                    }
                    else if (sheetName.Equals("NORM"))
                    {
                        cellCount = 9;
                    }
                    else if (sheetName.Equals("MAPs"))
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
                    if (j < 0)
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

        public static List<Dictionary<string, string>> STYLTableToData(DataTable STYLDt)
        {
            List<Dictionary<string, string>> STYLData = new List<Dictionary<string, string>>();
            Dictionary<string, string> colorDictionary = new Dictionary<string, string>();
            Dictionary<string, string> hatchDictionary = new Dictionary<string, string>();
            Dictionary<string, string> drawDictionary = new Dictionary<string, string>();
            Dictionary<string, string> stationDictionary = new Dictionary<string, string>();
            STYLData.Add(colorDictionary);
            STYLData.Add(hatchDictionary);
            STYLData.Add(drawDictionary);
            STYLData.Add(stationDictionary);
            int startRow = 1;
            string data = STYLDt.Rows[startRow][1].ToString();
            while (!(data.IndexOf("填充风格") > 0))
            {
                colorDictionary.Add(STYLDt.Rows[startRow]["ID"].ToString(), data);
                startRow++;
                data = STYLDt.Rows[startRow]["Item"].ToString();
            }
            startRow++;
            data = STYLDt.Rows[startRow]["Item"].ToString();
            while (!(data.IndexOf("绘图") > 0))
            {
                hatchDictionary.Add(STYLDt.Rows[startRow]["ID"].ToString(), data);
                startRow++;
                data = STYLDt.Rows[startRow]["Item"].ToString();
            }

            //drawId是绘图参数中据“绘图”一行的偏移量
            int drawId = 1;
            data = STYLDt.Rows[startRow + drawId]["Item"].ToString();
            while (!(data.IndexOf("指定电站") > 0))
            {
                drawDictionary.Add(drawId.ToString(), data);
                drawId++;
                data = STYLDt.Rows[startRow + drawId]["Item"].ToString();
                //吴老师在第九行的备注里加了要用的东西，虽然我记得说过备注里的东西用不上，只能打个补丁
                if (drawId == 9)
                {
                    data = STYLDt.Rows[startRow + drawId]["备注"].ToString();
                }
            }

            startRow += drawId;
            //总共有五个电站,因此只需要读取五个电站的数据
            for (int i = 1; i <= 5; i++)
            {
                //"*"用于将两个电站ID和电站名称分隔开
                stationDictionary.Add(i.ToString(), STYLDt.Rows[startRow + i]["Item"].ToString() + "*" + STYLDt.Rows[startRow + i]["备注"].ToString());

            }

            return STYLData;
        }

        public static List<Dictionary<string, Dictionary<string, string>>> NORMTableToData(DataTable NORMDt, int[] baseLoad)
        {
            List<Dictionary<string, Dictionary<string, string>>> NORMData = new List<Dictionary<string, Dictionary<string, string>>>();
            Dictionary<string, Dictionary<string, string>> drawParameter = new Dictionary<string, Dictionary<string, string>>();
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
                    dictionary.Add("secondMark", NORMDt.Rows[startRow]["Mark"].ToString());
                }
                else
                {
                    Dictionary<string, string> dictionary = new Dictionary<string, string>();
                    dictionary.Add("firstItem", NORMDt.Rows[startRow]["Item"].ToString());
                    dictionary.Add("firstHatch", NORMDt.Rows[startRow]["Hatch"].ToString());
                    dictionary.Add("firstARGB", NORMDt.Rows[startRow]["ARGB"].ToString());
                    dictionary.Add("firstMark", NORMDt.Rows[startRow]["Mark"].ToString());
                    LogoParameter.Add(NORMDt.Rows[startRow]["Flag"].ToString(), dictionary);
                }
                startRow++;
            }
            startRow++;
            for (; startRow < NORMDt.Rows.Count; startRow++)
            {
                Dictionary<string, string> dictionary = new Dictionary<string, string>();
                dictionary.Add("Hatch", NORMDt.Rows[startRow]["Hatch"].ToString());
                dictionary.Add("Item", NORMDt.Rows[startRow]["Item"].ToString());
                dictionary.Add("ARGB", NORMDt.Rows[startRow]["ARGB"].ToString());
                dictionary.Add("Mark", "1");
                if (!(drawParameter.ContainsKey(NORMDt.Rows[startRow]["Flag"].ToString())))
                {
                    drawParameter.Add(NORMDt.Rows[startRow]["Flag"].ToString(), dictionary);
                }

            }
            for (int i = 0; i < 366; i++)
            {
                if (i < 183)
                {
                    baseLoad[i] = int.Parse(NORMDt.Rows[i]["基荷A"].ToString());
                }
                else
                {
                    baseLoad[i] = int.Parse(NORMDt.Rows[i - 183]["基荷B"].ToString());
                }
            }
            return NORMData;
        }
        public static Dictionary<string, Dictionary<string, List<string>>> MAPsTableToData_Line(DataTable MAPsDt)
        {
            Dictionary<string, Dictionary<string, List<string>>> flagDictionary = new Dictionary<string, Dictionary<string, List<string>>>();
            int startRow = 0;
            for (; startRow < MAPsDt.Rows.Count; startRow++)
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
                for (int startLine = 1; startLine <= 24; startLine++)
                {
                    data.Add(MAPsDt.Rows[startRow][startLine].ToString());
                }
                dateDictionary.Add(MAPsDt.Rows[startRow]["Date"].ToString(), data);
            }
            return flagDictionary;
        }

        public static Dictionary<string, Dictionary<string, List<string>>> MAPsTableToData_Day(DataTable MAPsDt)
        {
            Dictionary<string, Dictionary<string, List<string>>> dateDictionary = new Dictionary<string, Dictionary<string, List<string>>>();
            int startRow = 0;
            for (; startRow < MAPsDt.Rows.Count; startRow++)
            {
                string date = MAPsDt.Rows[startRow]["Date"].ToString();
                Dictionary<string, List<string>> flagDictionary;
                List<string> data = new List<string>();
                if (!dateDictionary.ContainsKey(date))
                {
                    flagDictionary = new Dictionary<string, List<string>>();
                    dateDictionary.Add(date, flagDictionary);
                }
                else
                {
                    dateDictionary.TryGetValue(date, out flagDictionary);
                }
                for (int startLine = 1; startLine <= 24; startLine++)
                {
                    data.Add(MAPsDt.Rows[startRow][startLine].ToString());
                }
                flagDictionary.Add(MAPsDt.Rows[startRow]["Flag"].ToString(), data);
            }
            return dateDictionary;
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
        public float smallFontSize = 8.5F;
        public System.Drawing.Printing.PageSettings pageSettings { get; set; }
        public Rectangle drawArea { get; set; }
        public int logoWidth = 300;
        public bool drawed = false;
        public int maxRectangleY = 10000;
    }
}
