using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class StressTest
    {
        int work_face;
        //模板变量
        private Workbook workBook;
        private Worksheet workSheet;
        //报表变量
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;
        //报表路径
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        //连接数据库
        string constr = "server=192.168.0.101;database=20171113;uid=lalala;pwd=8886122";
        public StressTest(int work)
        {
            work_face = work;//工作面
            OpenExcel();
            //
            ConnectSQL(constr);
        }

        private void ConnectSQL(string constr)
        {
            string str = string.Format("SELECT DISTINCT Channel FROM MsgForewarn WHERE Location LIKE '{0}%' AND Location LIKE  '%皮带%' AND Location LIKE '%7m深%' AND CreatTime > '2018-2-17 16:00:00' AND CreatTime < '2018-2-18 16:00:00'", work_face);
            SqlConnection sqlConnection = new SqlConnection(constr);
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(str, sqlConnection);
            DataTable datatable = new DataTable();
            sqlDataAdapter.Fill(datatable);

            Calculate(sqlConnection, datatable);
        }

        private void OpenExcel()
        {
            string FilePath = @"C:\Users\14439\Desktop\yingpanhao\应力监测模板.xlsx";
            workBook = File.Exists(FilePath) ? new Workbook(FilePath) : new Workbook();
            workSheet = workBook.Worksheets[0];

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();

            if (workBook_excel.Worksheets["Sheet3"] != null)
                workSheet_excel = workBook_excel.Worksheets["Sheet3"];
            else
            {
                workBook_excel.Worksheets.Add("Sheet3");
                workSheet_excel = workBook_excel.Worksheets["Sheet3"];
            }

            workSheet_excel.Copy(workSheet);
            //CreateChart(workSheet_excel);
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        private void CreateChart(Worksheet workSheet_excel)
        {
            //B45 - K52
            //生成折线图
            workSheet_excel.Charts.Add(Aspose.Cells.Charts.ChartType.Line,45,2,52,11);
            Chart chart = workSheet_excel.Charts[0];
            chart.CategoryAxis.MajorGridLines.IsVisible = true;
            chart.CategoryAxis.MajorGridLines.Color = Color.Gray;

            //设置title样式
            chart.Title.Text = "胶运顺槽";
            chart.Title.TextFont.Color = Color.Gray;
            chart.Title.TextFont.IsBold = true;
            chart.Title.TextFont.Size = 12;

            chart.NSeries.Add("Sheet3!B2:E5",false);
            chart.NSeries.CategoryData = "Sheet3!B1:E1";
            Cells cells = workSheet_excel.Cells;
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                //设置每条折线的名称
                chart.NSeries[i].Name = cells[i + 1, 0].Value.ToString();

                //设置线的宽度
                chart.NSeries[i].Line.Weight = WeightType.MediumLine;

                //设置每个值坐标点的样式
                chart.NSeries[i].MarkerStyle = ChartMarkerType.Circle;
                chart.NSeries[i].MarkerSize = 5;
                chart.NSeries[i].MarkerBackgroundColor = Color.White;
                chart.NSeries[i].MarkerForegroundColor = Color.Gray;


                //每个折线向显示出值
                chart.NSeries[i].DataLabels.ShowValue = true;
                chart.NSeries[i].DataLabels.TextFont.Color = Color.Gray;

            }

            //设置x轴上数据的样式为灰色
            chart.CategoryAxis.TickLabels.Font.Color = Color.Gray;
            chart.CategoryAxis.TickLabelPosition = TickLabelPositionType.NextToAxis;

            //设置y轴的样式
            chart.ValueAxis.TickLabelPosition = TickLabelPositionType.Low;
            chart.ValueAxis.TickLabels.Font.Color = Color.Gray;
            // chart.ValueAxis.TickLabels.TextDirection = TextDirectionType.LeftToRight;
            //设置Legend位置以及样式
            chart.Legend.Position = LegendPositionType.Bottom;
            chart.Legend.TextFont.Color = Color.Gray;
            chart.Legend.Border.Color = Color.Gray;

            chart.Move(7,2,12,11);

        }

        private void Calculate(SqlConnection sqlConnection, DataTable datatable)
        {
            DataRow[] dataRow_channel = datatable.Select();
            double[] up_number = new double[dataRow_channel.Length];
            //dataRow[i]["Channel"].ToString()
            for (int i = 0; i < dataRow_channel.Length; i++)
            {
                string sqlstr = string.Format("SELECT Value FROM MsgForewarn WHERE Location LIKE '{0}%' AND Location LIKE  '%皮带%' AND Location LIKE '%7m深%' AND CreatTime > '2018-2-17 16:00:00' AND CreatTime < '2018-2-18 16:00:00' AND Channel = '{1}'", work_face, dataRow_channel[i]["Channel"].ToString());
               // string sqlstr = "SELECT Value FROM MsgForewarn WHERE Location LIKE '2101%' AND Location LIKE  '%皮带%' AND Location LIKE '%7m深%' AND CreatTime > '2018-2-17 16:00:00' AND CreatTime < '2018-2-18 16:00:00' AND Channel = '" + dataRow_channel[i]["Channel"].ToString() + "'";
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlstr,sqlConnection);
                DataTable datatable_1 = new DataTable();
                sqlDataAdapter.Fill(datatable_1);
                DataRow[] dr = datatable_1.Select();
                int len = dr.Length;

                up_number[i] = double.Parse(dr[len - 1]["Value"].ToString()) - double.Parse(dr[0]["Value"].ToString());
            }
            ToExcel();
        }

        private void ToExcel()
        {
            
        }

        public void AAA(double[] shuzu) {
            double a;
            for (int i = 0; i < shuzu.Length; i++)
            {
                a = shuzu[i];
            }
        }
    }
}
/**/