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

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Test
{
    public partial class 画图 : Form
    {
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
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.保存图片ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(139, 28);
            // 
            // 保存图片ToolStripMenuItem
            // 
            this.保存图片ToolStripMenuItem.Name = "保存图片ToolStripMenuItem";
            this.保存图片ToolStripMenuItem.Size = new System.Drawing.Size(138, 24);
            this.保存图片ToolStripMenuItem.Text = "保存图片";
            // 
            // 画图
            // 
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(1032, 683);
            this.Name = "画图";
            this.Text = "电站位置图";
            this.Load += new System.EventHandler(this.画图_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.画图_Paint);
            this.contextMenuStrip1.ResumeLayout(false);
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
            paintDays(g, 2020, 8, 30, 2, 80000);
        }

        // 开始绘图的年、月、日、绘图天数、一年最大值 
        private void paintDays(Graphics g,int year, int startMonth,int startDay, int days, int maxVal)
        {
            // 左右边界的宽度
            int leftBorderWidth = 80;
            int rightBorderWidth = 15;
            int topBorderWidth = 20;
            int bottomBorderWidth = 220;

            // 整个画面的宽度和高度
            int wholeWidth = 1020;
            int wholeHeight = 800;

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

            // 下面的均是测试代码，仅用于查看效果
            // 需要函数将数值转换为点坐标
            // 这里应该有一个循环，有多少种Flag，就循环多少次，每次画一种就行
            // 每种flag需要传来笔刷颜色，背景样式以及数据，数据用数组表示
            // 按理说这里的代码需要改动，需要调用解析文件的函数来获得具体的数据

            int types = 5;
            HatchStyle[] hatchStyles = new HatchStyle[types];
            // 这里是随机生成的填充样式
            for(int i = 0; i < types; i++)
            {
                switch (i % 3)
                {
                    case 0:
                        hatchStyles[i] = HatchStyle.Percent10;
                        break;
                    case 1:
                        hatchStyles[i] = HatchStyle.ForwardDiagonal;
                        break;
                    case 2:
                        hatchStyles[i] = HatchStyle.HorizontalBrick;
                        break;
                }    
            }
            // data是一种flag对应的数据
            int[] data = new int[] { 64727, 64726, 64338, 64938, 65239, 65428, 65723, 65723, 68403, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69506, 67437,
                                     64727, 64726, 64338, 64938, 65239, 65428, 65723, 65723, 68403, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69788, 69506, 67437 };

            // color是一种flag对应的颜色
            for (int i = 0; i < types; i++)
            {
                for(int j = 0; j < data.Length; j++)
                {
                    data[j] = data[j] - i * j * 80 - i * 56 - j * 28;
                }
                // data-数据 end-绘图的左上角坐标 width-绘图的宽度 height-绘图的高度 maxVal-一年的最大值，用于固定Y轴
                Point[] points = dataToPoint(data, end2, wholeWidth - leftBorderWidth - rightBorderWidth - 5, wholeHeight - topBorderWidth - bottomBorderWidth - 5, maxVal);
                // 笔刷中颜色、填充样式应该解析文件得到
                HatchBrush hBrush = new HatchBrush(hatchStyles[i], Color.Gray, Color.FromArgb(255, i * 579 % 255, i * 162 % 255, i * 561 % 255));
                // 画图
                g.FillPolygon(hBrush, points);
                g.DrawPolygon(Pens.Black, points);
            }

            


            // 绘制完图片部分，再去画坐标轴的间隔
            // 根据天数为1天还是多天，决定X轴分为24小时还是若干天
            int interval;
            if (days > 1)
            {
                interval = (end.X - start.X - 5) / (days);
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
    
        // 天数转换为日期
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

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            paintDays(g, 2020, 8, 30, 5, 80000);
        }

        private void pictureBox_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            
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
}
