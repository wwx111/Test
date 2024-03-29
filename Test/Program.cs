﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data;

namespace HUST_Grph
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>

        public static void EntryPoint()
        {
            Main(); // 调用原始的Main方法
        }

        [STAThread]
        static void Main()
        {
            //int year = 2020;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string path = "D:\\wu_picture_data\\picture_data23.12.06\\广西2030_2030_2_S0000.xlsx";
            //string path = "D:\\GTDS_0001\\GTDS_Exam_2030_2_S0000_new.xlsx";
            //string path = "D:\\input_11_2022020000_70_0.xlsx";
            //string path = "C:\\Users\\11852\\Desktop\\GTDS_Exam_2030_2_S0001.xlsx";
            DataSet ds = new DataSet();
            DataTable STYLdata = loadData.ExcelToDatatable(path, "STYL");
            DataTable NORMdata = loadData.ExcelToDatatable(path, "NORM");
            DataTable Mapsdata = loadData.ExcelToDatatable(path, "MAPs");
            STYLdata.TableName = "STYL";
            NORMdata.TableName = "NORM";
            Mapsdata.TableName = "MAPs";

            ds.Tables.Add(STYLdata);
            ds.Tables.Add(NORMdata);
            ds.Tables.Add(Mapsdata);


            画图 ht = new 画图(ds);
            ht.newTab();

            ht.Show();
            Application.Run(ht);


            Boolean tf0 = false;
            String filePath = @"D:\HUST_Grph\what.jpg";
            画图 directDraw;
            bool tf = (directDraw = new 画图(ds, tf0, filePath)) != null;


            /*loadData LoadData = new loadData();
            //string path = "C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-06\\input_11_2022020000_70_0.xlsx";
            //string path = "C:\\power_system\\data\\804\\XML\\input_11_2022020000_70_0.xlsx";
            //DataTable STYLdata = loadData.ExcelToDatatable(path, "STYL");
            DataTable NORMdata = loadData.ExcelToDatatable(path, "NORM");
            DataTable Mapsdata = loadData.ExcelToDatatable(path, "MAPs");
            //DataTable data = loadData.ExcelToDatatable("C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-05\\HUST_ViewConfig.xlsx", "MAPs");
            List<Dictionary<string, string>> STYLlist = loadData.STYLTableToData(STYLdata);
            List<Dictionary<string, Dictionary<string, string>>> NORMlist = loadData.NORMTableToData(NORMdata);
            Dictionary<string, Dictionary<string, List<string>>> dictionary = loadData.MAPsTableToData(Mapsdata);*/
        }
    }
}
