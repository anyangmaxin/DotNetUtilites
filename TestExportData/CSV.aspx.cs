using System;
using System.Data;
using ExportDataToXXX;

namespace TestExportData
{
    public partial class CSV : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }


  

        protected void OnServerClick(object sender, EventArgs e)
        {
            DataTable dt=new DataTable();
            dt.Columns.Add("A",typeof(string));
            dt.Columns.Add("1", typeof(string));
            dt.Columns.Add("中文", typeof(string));

            for (int i = 0; i < 100; i++)
            {
                DataRow newRow = dt.NewRow();
                newRow["A"] = "测试数据A" + i;
                newRow["1"] = "测试数据1" + i;
                newRow["中文"] = "测试数据中文" + i;

                dt.Rows.Add(newRow);
            }
            ExportToCSV exportToCsv=new ExportToCSV(dt);
            exportToCsv.DataToCSV("测试");


            btnTest.Value = "已经改变";
        }
    }
}