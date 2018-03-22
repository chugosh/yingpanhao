using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class DrillingTest
    {
        //数据库服务器
        string constr = "server=.;database=UPRESSURE;uid=sa;pwd=sdkjdx";
        //模板excel文件
        private Workbook workBook;
        //模板工作sheet
        private Worksheet workSheet;
        //报表excel 钻屑法sheet
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;
        //总报表excel 单例模式
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        //查询语句
        static string sqlquery = "SELECT 单孔平均,单孔最大值,最大孔深,距工作面距离 FROM 钻屑法数据表 ";
        string sqlquery_jiaoyun = string.Format(sqlquery + "WHERE 日期 = '{0}' AND 胶运 = 1 ORDER BY 单孔最大值 DESC", new GetTime().getDateToday());
        string sqlquery_fuyun = string.Format(sqlquery + "WHERE 日期 = '{0}' AND 胶运 = 0 ORDER BY 单孔最大值 DESC", new GetTime().getDateToday());
        //数据库连接
        SqlConnection sqlConnection;
        SqlDataAdapter sqlDataAdapter;
        //胶运数据表 辅运数据表
        DataTable dataTable_j;
        DataTable dataTable_f;
        public void Start()
        {
            //打开excel
            OpenExcel();

            ConnectSQL();
            dataTable_j = GetDataTable(sqlquery_jiaoyun);
            dataTable_f = GetDataTable(sqlquery_fuyun);
            //计算胶运 辅运
            Calculate_jiaoyun(dataTable_j);
            Calculate_fuyun(dataTable_f);

            SaveReportFile();
        }

        private void ConnectSQL()
        {
            sqlConnection = new SqlConnection(constr);
        }

        private void Calculate_fuyun(DataTable datatable)
        {
            DataRow[] dataRow = datatable.Select();
            int number = dataRow.Length;
            double max = double.Parse(dataRow[0]["单孔最大值"].ToString());
            int maxDeep = int.Parse(dataRow[0]["最大孔深"].ToString());
            int working_face_position = int.Parse(dataRow[0]["距工作面距离"].ToString());
            double maxAverage = double.Parse(dataRow[0]["单孔平均"].ToString());
            for (int i = 1; i < dataRow.Length; i++)
            {
                if (double.Parse(dataRow[i]["单孔平均"].ToString()) > maxAverage)
                {
                    maxAverage = double.Parse(dataRow[i]["单孔平均"].ToString());
                }
            }
            InsertValue_fuyun(number, max, maxDeep, maxAverage, working_face_position);
        }

        private void Calculate_jiaoyun(DataTable datatable)
        {
            DataRow[] dataRow         = datatable.Select();
            int number                = dataRow.Length;
            double max                = double.Parse(dataRow[0]["单孔最大值"].ToString());
            int maxDeep               = int.Parse(dataRow[0]["最大孔深"].ToString());
            int working_face_position = int.Parse(dataRow[0]["距工作面距离"].ToString());
            double maxAverage         = double.Parse(dataRow[0]["单孔平均"].ToString());
            for (int i = 1; i < dataRow.Length; i++)
            {
                if (double.Parse(dataRow[i]["单孔平均"].ToString()) > maxAverage)
                {
                    maxAverage = double.Parse(dataRow[i]["单孔平均"].ToString());
                }
            }
            InsertValue_jiaoyun(number, max, maxDeep, maxAverage, working_face_position);
        }

        private DataTable GetDataTable(string str)
        {
            sqlDataAdapter = new SqlDataAdapter(str, sqlConnection);
            DataTable datatable = new DataTable();
            sqlDataAdapter.Fill(datatable);
            return datatable;

        }

        private void OpenExcel()
        {
            //找到模板
            string FilePath = @"C:\Users\14439\Desktop\yingpanhao\钻屑法模板.xlsx";
            //存在就创建
            workBook = File.Exists(FilePath) ? new Workbook(FilePath) : new Workbook();
            workSheet = workBook.Worksheets[0];

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();
            workSheet_excel = workBook_excel.Worksheets[0];
            //将模板拷贝到报表excel
            workSheet_excel.Copy(workSheet);
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        //辅运通道
        private void InsertValue_fuyun(int number, double max, int maxDeep, double maxAverage, int working_face_position)
        {
            Cell cellItem_number = workSheet_excel.Cells["E3"];
            cellItem_number.PutValue(number);
            Cell cellItem1 = workSheet_excel.Cells["F3"];
            cellItem1.PutValue(maxAverage);
            Cell cellItem2 = workSheet_excel.Cells["J3"];
            cellItem2.PutValue(max);
            Cell cellItem3 = workSheet_excel.Cells["M3"];
            cellItem3.PutValue(maxDeep);
            Cell cellItem6 = workSheet_excel.Cells["H3"];
            cellItem6.PutValue("距工作面" + working_face_position + "米!!!");
            if (maxDeep <= 10 && maxDeep >= 1)
            {
                Cell cellItem4 = workSheet_excel.Cells["O3"];
                if (max > 3.5)
                {
                    cellItem4.PutValue("是");
                    Cell cellItem5 = workSheet_excel.Cells["R3"];
                    cellItem5.PutValue("有冲击危险");
                }
                else
                {
                    cellItem4.PutValue("否");
                    Cell cellItem5 = workSheet_excel.Cells["R3"];
                    cellItem5.PutValue("无冲击危险");
                }

            }

        }
        //胶运通道
        private void InsertValue_jiaoyun(int number, double max, int maxDeep, double maxAverage, int working_face_position)
        {
            Cell cellItem_number = workSheet_excel.Cells["E4"];
            cellItem_number.PutValue(number);
            Cell cellItem1 = workSheet_excel.Cells["F4"];
            cellItem1.PutValue(maxAverage);
            Cell cellItem2 = workSheet_excel.Cells["J4"];
            cellItem2.PutValue(max);
            Cell cellItem3 = workSheet_excel.Cells["M4"];
            cellItem3.PutValue(maxDeep);
            Cell cellItem6 = workSheet_excel.Cells["H4"];
            cellItem6.PutValue("距工作面" + working_face_position + "米!!!");
            if (maxDeep <= 10 && maxDeep >= 1)
            {
                Cell cellItem4 = workSheet_excel.Cells["O4"];
                if (max > 3.5)
                {
                    cellItem4.PutValue("是");
                    Cell cellItem5 = workSheet_excel.Cells["R4"];
                    cellItem5.PutValue("有冲击危险");
                }
                else
                {
                    cellItem4.PutValue("否");
                    Cell cellItem5 = workSheet_excel.Cells["R4"];
                    cellItem5.PutValue("无冲击危险");
                }

            }

        }

        private void SaveReportFile()
        {
            if (!Directory.Exists(@"C:\Users\14439\Desktop\yingpanhao\报表"))
                Directory.CreateDirectory(@"C:\Users\14439\Desktop\yingpanhao\报表");
            //  设置执行公式计算 - 如果代码中用到公式，需要设置计算公式，导出的报表中，公式才会自动计算
            workBook_excel.CalculateFormula(true);
            //  保存文件
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

    }
}
