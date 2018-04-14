using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using Excel = Aspose.Cells;
using System.IO;
using System.Data;
using CCWin.SkinControl;

namespace WindowsFormsApplication1
{
    
    class AsposeTest
    {
        public AsposeTest()
        {
            //引入证书


            //string filePath = @"C:\Users\14439\Desktop\yingpanhao\报表\";
            //string ReportFileName = string.Format("Excel_{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd"));
            //Workbook wb = new Workbook();
            ////保存文件
            //wb.Save(filePath + ReportFileName, SaveFormat.Xlsx);

        }

        //public void Start()
        //{
        //    OpenExcel();
        //    InsertValue();
        //    SaveReportFile();
        //}

        //public void SetPicture(string filePath)
        //{
        //    OpenExcel();
        //    try
        //    {
        //        workSheet.Pictures.Add(29,2,filePath);
        //        SaveReportFile();
        //    }
        //    catch (Exception){
        //    }


        //}

        //private void InsertValue()
        //{
        //    InsertValue_fuyun(number, max, maxDeep, singleAverage, working_face_position);
        //    InsertValue_jiaoyun(number, max, maxDeep, singleAverage, working_face_position);
        //    //Insert(number, max, maxDeep, singleAverage, working_face_position);
        //}

        ////辅运通道
        //private void InsertValue_fuyun(int number,double[] max, int[] maxDeep, double[] singleAverage,int[] working_face_position)
        //{
        //    int location = 0;
        //    double maxAverage = singleAverage[0];
        //    double maxValue = max[0];

        //    Cell cellItem_number = workSheet.Cells["E23"];
        //    cellItem_number.PutValue(number);
        //    //singAverage[]
        //    for (int i = 1;i < singleAverage.Length;i++)
        //    {
        //        if (singleAverage[i] > maxAverage)
        //            maxAverage = singleAverage[i];
        //    }
        //    Cell cellItem1 = workSheet.Cells["F23"];
        //    cellItem1.PutValue(maxAverage);

        //    //max[]
        //    for (int i = 1;i < max.Length;i++)
        //    {
        //        if (max[i] > maxValue)
        //        {
        //            maxValue = max[i];
        //            location = i;
        //        }
        //    }
        //    //
        //    Cell cellItem2 = workSheet.Cells["J23"];
        //    cellItem2.PutValue(maxValue);
        //    Cell cellItem3 = workSheet.Cells["M23"];
        //    cellItem3.PutValue(maxDeep[location]);
        //    Cell cellItem6 = workSheet.Cells["H23"];
        //    cellItem6.PutValue("距工作面" + working_face_position[location] + "米");

        //    if (maxDeep[location] <= 10 && maxDeep[location] >=1)
        //    {
        //        Cell cellItem4 = workSheet.Cells["O23"];
        //        if (maxValue > 3.5)
        //        {
        //            cellItem4.PutValue("是");
        //            Cell cellItem5 = workSheet.Cells["R23"];
        //            cellItem5.PutValue("有冲击危险");
        //        }
        //        else
        //        {
        //            cellItem4.PutValue("否");
        //            Cell cellItem5 = workSheet.Cells["R23"];
        //            cellItem5.PutValue("无冲击危险");
        //        }
                    
        //    }

        //}
        ////胶运通道
        //private void InsertValue_jiaoyun(int number, double[] max, int[] maxDeep, double[] singleAverage, int[] working_face_position)
        //{
        //    int location = 0;
        //    double maxAverage = 0;
        //    double maxValue = 0;
        //    Cell cellItem_number = workSheet.Cells["E24"];
        //    cellItem_number.PutValue(number);
        //    //singAverage[]
        //    for (int i = 1; i < singleAverage.Length; i++)
        //    {
        //        maxAverage = singleAverage[0];
        //        if (singleAverage[i] > maxAverage)
        //            maxAverage = singleAverage[i];
        //    }
        //    Cell cellItem1 = workSheet.Cells["F24"];
        //    cellItem1.PutValue(maxAverage);

        //    //max[]
        //    for (int i = 1; i < max.Length; i++)
        //    {
        //        maxValue = max[0];
        //        if (maxValue < max[i])
        //        {
        //            maxValue = max[i];
        //            location = i;
        //        }
        //    }
        //    //
        //    Cell cellItem2 = workSheet.Cells["J24"];
        //    cellItem2.PutValue(maxValue);
        //    Cell cellItem3 = workSheet.Cells["M24"];
        //    cellItem3.PutValue(maxDeep[location]);
        //    Cell cellItem6 = workSheet.Cells["H24"];
        //    cellItem6.PutValue("距工作面" + working_face_position[location] + "米");

        //    if (maxDeep[location] <= 10 && maxDeep[location] >= 1)
        //    {
        //        Cell cellItem4 = workSheet.Cells["O24"];
        //        if (maxValue > 3.5)
        //        {
        //            cellItem4.PutValue("是");
        //            Cell cellItem5 = workSheet.Cells["R24"];
        //            cellItem5.PutValue("有冲击危险");
        //        }
        //        else
        //        {
        //            cellItem4.PutValue("否");
        //            Cell cellItem5 = workSheet.Cells["R24"];
        //            cellItem5.PutValue("无冲击危险");
        //        }
                    
        //    }

        //}

        //private void OpenExcel()
        //{
        //    string FilePath = @"C:\Users\14439\Desktop\yingpanhao\营盘壕煤矿2101工作面矿压分析综合日报表3.2.xlsx";
        //    workBook = File.Exists(FilePath) ? new Workbook(FilePath) : new Workbook();
        //    workSheet = workBook.Worksheets[0];
        //}

        //private void SaveReportFile()
        //{
        //    if (!Directory.Exists(@"C:\Users\14439\Desktop\yingpanhao"))
        //        Directory.CreateDirectory(@"C:\Users\14439\Desktop\yingpanhao");

        //    //  设置执行公式计算 - 如果代码中用到公式，需要设置计算公式，导出的报表中，公式才会自动计算
        //    workBook.CalculateFormula(true);

        //    //  生成的文件名称
        //    string ReportFileName = string.Format("Excel_{0}.xlsx", DateTime.Now.ToString("yyyy-MM-dd"));

        //    //  保存文件
        //    workBook.Save(@"C:\Users\14439\Desktop\yingpanhao\" + ReportFileName, SaveFormat.Xlsx);
        //}


        //        //datatable
        //private void Insert(int number, double[] max, int[] maxDeep, double[] singleAverage, int[] working_face_position)
        //{
        //    DataTable dataTable = GetData( number, max,  maxDeep,  singleAverage, working_face_position);
        //    //  写入数据的起始位置
        //    string cell_start_region = "E3";
        //    //  获得开始位置的行号
        //    int startRow = workSheet.Cells[cell_start_region].Row;
        //    //  获得开始位置的列号  
        //    int startColumn = workSheet.Cells[cell_start_region].Column;

        //    //  写入Excel。参数说明，直接查阅文档
        //    workSheet.Cells.ImportDataTable(dataTable, false, startRow, startColumn, false, true);
        //}
        //private DataTable GetData(int number, double[] max, int[] maxDeep, double[] singleAverage, int[] working_face_position)
        //{
        //    DataTable dataTable = new DataTable();
        //    dataTable.Columns.Add("number", typeof(int));
        //    dataTable.Columns.Add("average", typeof(double));
        //    dataTable.Columns.Add("work_position", typeof(string));
        //    dataTable.Columns.Add("meifenliang", typeof(double));
        //    dataTable.Columns.Add("deep", typeof(int));
        //    dataTable.Columns.Add("chaobiao", typeof(string));
        //    dataTable.Columns.Add("comment", typeof(string));
        //    //int a = 7;
        //    dataTable.Rows.Add(new object[] { 1, 2.16, "mianjisss", 2.66, 3, "sdfsd", "sadfasdf" });
        //    //dataTable.Rows.Add(new object[] { 2, 2.46, "mianji" + a.ToString(), 2.86, 9, "qwer", "zzxzvbcv" });

        //    return dataTable;
        //}

        //private void GetMerge()
        //{
        //    //Range range = new Cell().GetMergedRange();
        //}
       
    }
}
