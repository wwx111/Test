using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data;

namespace Test
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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            画图 ht = new 画图();
            ht.newTab(2020);
            Application.Run(ht);

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
