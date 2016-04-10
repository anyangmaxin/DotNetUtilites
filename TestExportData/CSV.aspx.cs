﻿using System;
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
            dt.Columns.Add("A,",typeof(string));
            dt.Columns.Add("1", typeof(string));
            dt.Columns.Add("中,文，", typeof(string));

            for (int i = 0; i < 100; i++)
            {
                DataRow newRow = dt.NewRow();
                newRow["A,"] = "测试，数据A" + i;
                newRow["1"] = "测试,数据1" + i;
                newRow["中,文，"] = "测试,，数据中文" + i;
                dt.Rows.Add(newRow);
            }
            ExportToCSV exportToCsv=new ExportToCSV(dt);
            exportToCsv.DataToCSV("测试只传入dt");
        }

        protected void btnTest2_OnServerClick(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            //dt.Columns.Add("A", typeof(string));
            //dt.Columns.Add("1", typeof(string));
            //dt.Columns.Add("中文", typeof(string));

            for (int i = 0; i < 100; i++)
            {
                DataRow newRow = dt.NewRow();
                newRow[0] = "测试数据A" + i;
                newRow[1] = "测试数据1" + i;
                newRow[2] = "测试数据中文" + i;
                dt.Rows.Add(newRow);
            }
            ExportToCSV exportToCsv = new ExportToCSV(new[] { "第一列A", "第二列1", "第三列" }, dt);
            exportToCsv.DataToCSV("测试传入标题数组以及dt");
        }

        protected void btnTest3_OnServerClick(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("A", typeof (string));
            dt.Columns.Add("1", typeof (string));
            dt.Columns.Add("中文", typeof (string));

            for (int i = 0; i < 3; i++)
            {
                DataRow newRow = dt.NewRow();
                newRow["A"] = "测试数据A" + i;
                newRow["1"] = "测试数据1" + i;
                newRow["中文"] = "测试数据中文" + i;
                dt.Rows.Add(newRow);
            }
            ExportToCSV exportToCsv = new ExportToCSV(new[] {"第一列A", "第二列1", "第三列"}, new[,]{
            {"1第一个字段1", "1第二个字段A", "1第三个字段"}, 
            {"2第一个字段1", "2第二个字段A", "2第三个字段"},
                
            {
                "3第一个字段1",
                "3第二个字段A",
                "3第三个字段"
            }
        }
    , dt);
            exportToCsv.DataToCSV("测试传入标题数组以及dt以及字段名称数组");
            
        }
    }
}