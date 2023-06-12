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
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new 画图());
            loadData LoadData = new loadData();
            DataTable STYLdata = loadData.ExcelToDatatable("C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-06\\input_11_2022020000_70_0.xlsx", "STYL");
            DataTable NORMdata = loadData.ExcelToDatatable("C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-06\\input_11_2022020000_70_0.xlsx", "NORM");
            DataTable Mapsdata = loadData.ExcelToDatatable("C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-06\\input_11_2022020000_70_0.xlsx", "MAPs");
            //DataTable data = loadData.ExcelToDatatable("C:\\Users\\11852\\Documents\\WeChat Files\\wxid_wk1qwav6tqmv12\\FileStorage\\File\\2023-05\\HUST_ViewConfig.xlsx", "MAPs");
            List<Dictionary<string, string>> STYLlist = loadData.STYLTableToData(STYLdata);
            List<Dictionary<string, Dictionary<string, string>>> NORMlist = loadData.NORMTableToData(NORMdata);
            Dictionary<string, Dictionary<string, List<string>>> dictionary = loadData.MAPsTableToData(Mapsdata);
        }
    }
}
