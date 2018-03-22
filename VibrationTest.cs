using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsFormsApplication1
{
    class VibrationTest
    {
        string[] column = { "Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9", "Column10" };
        string excelFilePath = CreateExcelTest.GetCreateExcelTest().ExcelFilePath();
        //模板变量
        private Workbook workBook;
        private Worksheet workSheet;
        //报表变量
        private Workbook workBook_excel;
        private Worksheet workSheet_excel;

        private int times_all,times_front, times_behind;//总次数

        public int Times_all
        {
            get
            {
                return times_all;
            }

            set
            {
                times_all = value;
            }
        }

        public int Times_front
        {
            get
            {
                return times_front;
            }

            set
            {
                times_front = value;
            }
        }

        public int Times_behind
        {
            get
            {
                return times_behind;
            }

            set
            {
                times_behind = value;
            }
        }

        public VibrationTest()
        {
            OpenExcel();
        }

        //添加图片
        public void AddPictures(string filePath_1,string filePath_2)
        {
            //添加图片
            if (filePath_2 == null)
            {
                //左行左列有行右列 + 文件
                workSheet_excel.Pictures.Add(5, 2, 16, 11, filePath_1);
                SaveReportFile();
            }
            if (filePath_1 == null)
            {
                //L5 T12
                workSheet_excel.Pictures.Add(5, 12, 12, 20, filePath_2);
                SaveReportFile();
            }
            
            
        }

        //打开报表模板sheet2
        private void OpenExcel()
        {
            string FilePath = @"C:\Users\14439\Desktop\yingpanhao\微震模板.xlsx";
            workBook = File.Exists(FilePath) ? new Workbook(FilePath) : new Workbook();
            workSheet = workBook.Worksheets[0];

            workBook_excel = CreateExcelTest.GetCreateExcelTest().GetWorkBookExcel();

            if (workBook_excel.Worksheets["Sheet2"] != null )
                workSheet_excel = workBook_excel.Worksheets["Sheet2"];
            else
            {
                workBook_excel.Worksheets.Add("Sheet2");
                workSheet_excel = workBook_excel.Worksheets["Sheet2"];
            }

            workSheet_excel.Copy(workSheet);
            workBook_excel.Save(excelFilePath, SaveFormat.Xlsx);
        }

        //保存文件
        private void SaveReportFile()
        {
            if (!Directory.Exists(@"C:\Users\14439\Desktop\yingpanhao\报表"))
                Directory.CreateDirectory(@"C:\Users\14439\Desktop\yingpanhao\报表");
            //  设置执行公式计算 - 如果代码中用到公式，需要设置计算公式，导出的报表中，公式才会自动计算
            workBook_excel.CalculateFormula(true);
            //  保存文件
            workBook_excel.Save(excelFilePath);
        }

        //excel to datatable
        public void GetDataTable()
        {
            //微震数据从excel中导入到的DataTable
            string excel_to_datatable_file_path = @"C:\Users\14439\Desktop\yingpanhao\微震报表.xlsx";
            //微震表的数据excel
            Workbook wb;
            wb = File.Exists(excel_to_datatable_file_path) ? new Workbook(excel_to_datatable_file_path) : new Workbook();
            Worksheet ws = wb.Worksheets[0];
            DataTable datatable = ws.Cells.ExportDataTable(1, 0, ws.Cells.MaxRow+1, ws.Cells.MaxColumn+1);

            CalculateAll(datatable);

        }

        //计算数据
        private void CalculateAll(DataTable datatable)
        {
            DataRow[] dataRow_front = datatable.Select("column10 LIKE '%面前%'");
            DataRow[] dataRow_behind = datatable.Select("column10 LIKE '%面后%'");
            
            Times_behind = dataRow_behind.Length;//面后次数
            Times_front = dataRow_front.Length;//面前次数
            Times_all = Times_front + Times_behind;//总次数
            
            double sum_energy_front = 0;//面前总能量
            double sum_energy_behind = 0; //面后总能量
            double max_energy, quake_level;//最大能量 震级
            double sum_energy = 0;
            int sum = 0; 
            string location;//最大能量的位置（备注）

            DataView dataView = datatable.DefaultView;//
            dataView.Sort = "Column6 DESC";//按照第六列数字大小排序,从大到小
            DataTable datatable_sort = dataView.ToTable();//转换成新的排好序的数据表
            //每一行存储到dataRow数组里 已经按照从大到小的顺序存储 最大的都在第一行！！！
            DataRow[] dataRows = datatable_sort.Select("Column10 LIKE '%面%'");

            //找到最大能量 算出震级 找到相应备注的位置
            max_energy = System.Convert.ToDouble(dataRows[0][column[5]].ToString());
            quake_level = (Math.Log10(max_energy) - 1.8) / 1.9;
            location = dataRows[0][column[9]].ToString();

            //算出 面前 和 面后 的总能量
            for (int i = 0; i < dataRow_front.Length; i++)
                sum_energy_front += System.Convert.ToDouble(dataRow_front[i][column[5]].ToString());
            for (int i = 0; i < dataRow_behind.Length; i++)
                sum_energy_behind += System.Convert.ToDouble(dataRow_behind[i][column[5]].ToString());

            for (int i = 0; i < dataRows.Length; i++)
            {
                if (System.Convert.ToDouble(dataRows[i][column[5]].ToString()) > 20000)
                {
                    sum_energy += System.Convert.ToDouble(dataRows[i][column[5]].ToString());
                    sum++;
                }
            }

            string time = dataRows[0][column[1]].ToString().Substring(6)
                            + " " + dataRows[0][column[1]].ToString().Substring(0, 5);
            //写入excel
            ToExcel(sum_energy_front, sum_energy_behind, max_energy, location, quake_level,time,sum_energy,sum);
        }

        //导入excel
        private void ToExcel(double sum_energy_front, 
                                double sum_energy_behind,
                                double max_energy,
                                string location,
                                double quake_level,
                                string time,
                                double sum_energy,
                                int sum)
        {
            //L3,O3
            Cell cellItem_times_all = workSheet_excel.Cells["B3"];
            Cell cellItem_times_front = workSheet_excel.Cells["D3"];
            Cell cellItem_times_behind = workSheet_excel.Cells["G3"];
            cellItem_times_all.PutValue(Times_all);
            cellItem_times_front.PutValue(Times_front);
            cellItem_times_behind.PutValue(Times_behind);

            Cell cellItem_sum_energy_front = workSheet_excel.Cells["E3"];
            cellItem_sum_energy_front.PutValue(sum_energy_front);
            Cell cellItem_sum_energy_behind = workSheet_excel.Cells["H3"];
            cellItem_sum_energy_behind.PutValue(sum_energy_behind);
            Cell cellItem_sum_energy = workSheet_excel.Cells["C3"];
            cellItem_sum_energy.PutValue(sum_energy_behind + sum_energy_front);

            Cell cellItem_max_energy = workSheet_excel.Cells["L3"];
            Cell cellItem_location = workSheet_excel.Cells["O3"];
            Cell cellItem_level = workSheet_excel.Cells["N3"];
            cellItem_max_energy.PutValue(max_energy);
            cellItem_location.PutValue(location);
            cellItem_level.PutValue(quake_level);

            Cell cellItem_time = workSheet_excel.Cells["J3"];
            cellItem_time.PutValue(time);
            //Q3 S3
            Cell cellItem_sum_energy_2 = workSheet_excel.Cells["Q3"];
            cellItem_sum_energy_2.PutValue(sum_energy);
            Cell cellItem_sum = workSheet_excel.Cells["S3"];
            cellItem_sum.PutValue(sum);

            SaveReportFile();
        }
    }
}



//foreach (DataRow dr in dataRow_front)
//{
//    sum_energy_front += System.Convert.ToDouble(dr["Column6"].ToString());
//}

//foreach (DataRow dr in dataRow_behind)
//{
//    sum_energy_behind += System.Convert.ToDouble(dr["Column6"].ToString());
//}

/*****计算面前面后最大值********/

//for (int i = 0; i < dataRow_front.Length; i++)
//{
//    energy_front[i] = System.Convert.ToDouble(dataRow_front[i]["Column6"].ToString());
//    s_front[i] = dataRow_front[i]["Column10"].ToString();
//}

//for (int i = 0; i < dataRow_behind.Length; i++)
//{
//    energy_behind[i] = System.Convert.ToDouble(dataRow_behind[i]["Column6"].ToString());
//    s_behind[i] = dataRow_behind[i]["Column10"].ToString();
//}