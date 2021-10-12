using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WIMDataProcessingApp
{
    /// <summary>
    /// 每日交通流量，对应“车流量统计表”
    /// </summary>
    public class DailyTraffic
    {
        public string Date { get; set; }
        public int UpStreamCount { get; set; }

        public int DownStreamCount { get; set; }

        public int TotalStreamCount { get; set; }
    }
    public class HighSpeedData
    {
        public int HSData_Id { get; set; }
        public Nullable<byte> Lane_Id { get; set; }
        public Nullable<System.DateTime> HSData_DT { get; set; }
        public string Oper_Direc { get; set; }
        public Nullable<byte> Axle_Num { get; set; }
        public Nullable<byte> AxleGrp_Num { get; set; }
        public Nullable<int> Gross_Load { get; set; }
        public Nullable<short> Veh_Type { get; set; }
        public Nullable<int> LWheel_1_W { get; set; }
        public Nullable<int> LWheel_2_W { get; set; }
        public Nullable<int> LWheel_3_W { get; set; }
        public Nullable<int> LWheel_4_W { get; set; }
        public Nullable<int> LWheel_5_W { get; set; }
        public Nullable<int> LWheel_6_W { get; set; }
        public Nullable<int> LWheel_7_W { get; set; }
        public Nullable<int> LWheel_8_W { get; set; }
        public Nullable<int> RWheel_1_W { get; set; }
        public Nullable<int> RWheel_2_W { get; set; }
        public Nullable<int> RWheel_3_W { get; set; }
        public Nullable<int> RWheel_4_W { get; set; }
        public Nullable<int> RWheel_5_W { get; set; }
        public Nullable<int> RWheel_6_W { get; set; }
        public Nullable<int> RWheel_7_W { get; set; }
        public Nullable<int> RWheel_8_W { get; set; }
        public Nullable<int> AxleDis1 { get; set; }
        public Nullable<int> AxleDis2 { get; set; }
        public Nullable<int> AxleDis3 { get; set; }
        public Nullable<int> AxleDis4 { get; set; }
        public Nullable<int> AxleDis5 { get; set; }
        public Nullable<int> AxleDis6 { get; set; }
        public Nullable<int> AxleDis7 { get; set; }
        public Nullable<int> Violation_Id { get; set; }
        public Nullable<byte> OverLoad_Sign { get; set; }
        public Nullable<int> Speed { get; set; }
        public Nullable<decimal> Acceleration { get; set; }
        public Nullable<int> Veh_Length { get; set; }
        public Nullable<decimal> QAT { get; set; }
        public string License_Plate { get; set; }
        public string License_Plate_Color { get; set; }
        public string F7Code { get; set; }
        public Nullable<float> ExternInfo { get; set; }
        public Nullable<int> Temp { get; set; }
        public Nullable<int> SiteID { get; set; }
    }

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Calc_Click(object sender, RoutedEventArgs e)
        {

            var startDateTime = StartDateTime.SelectedDate ?? DateTime.Now.AddDays(-1);
            var finishDateTime = FinishDateTime.SelectedDate ?? DateTime.Now;
            //var StartDateTime = new DateTime(2020, 10, 1, 0, 0, 0);
            //var FinishDateTime = new DateTime(2020, 11, 1, 0, 0, 0);
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDateTime && x.HSData_DT < finishDateTime;

            using (var db = new HighSpeed_PROCEntities())
            {
                #region HS_DataForAnalysis
                var highSpeedData = (
                    from e1 in db.HS_Data_PROC
                    select new HighSpeedData
                    {
                        Acceleration = e1.Acceleration,
                        AxleGrp_Num = e1.AxleGrp_Num,
                        Axle_Num = e1.Axle_Num,
                        ExternInfo = e1.ExternInfo,
                        F7Code = e1.F7Code,
                        HSData_Id = e1.HSData_Id,
                        Veh_Length = e1.Veh_Length,
                        Veh_Type = e1.Veh_Type,
                        LWheel_1_W = e1.LWheel_1_W,
                        LWheel_2_W = e1.LWheel_2_W,
                        LWheel_3_W = e1.LWheel_3_W,
                        LWheel_4_W = e1.LWheel_4_W,
                        LWheel_5_W = e1.LWheel_5_W,
                        LWheel_6_W = e1.LWheel_6_W,
                        LWheel_7_W = e1.LWheel_7_W,
                        LWheel_8_W = e1.LWheel_8_W,
                        Lane_Id = e1.Lane_Id,
                        Oper_Direc = e1.Oper_Direc,
                        Speed = e1.Speed,
                        OverLoad_Sign = e1.OverLoad_Sign,
                        RWheel_1_W = e1.RWheel_1_W,
                        RWheel_2_W = e1.RWheel_2_W,
                        RWheel_3_W = e1.RWheel_3_W,
                        RWheel_4_W = e1.RWheel_4_W,
                        RWheel_5_W = e1.RWheel_5_W,
                        RWheel_6_W = e1.RWheel_6_W,
                        RWheel_7_W = e1.RWheel_7_W,
                        RWheel_8_W = e1.RWheel_8_W,
                        Violation_Id = e1.Violation_Id,
                        AxleDis1 = e1.AxleDis1,
                        AxleDis2 = e1.AxleDis2,
                        AxleDis3 = e1.AxleDis3,
                        AxleDis4 = e1.AxleDis4,
                        AxleDis5 = e1.AxleDis5,
                        AxleDis6 = e1.AxleDis6,
                        AxleDis7 = e1.AxleDis7,
                        HSData_DT = e1.HSData_DT,
                        Gross_Load = e1.Gross_Load,
                        License_Plate = e1.License_Plate
                    }
                    );
                #endregion;

                string SummaryString = string.Empty;

                //不同车道车数量分布
                //var Lane_Div = new int[] { 1, 2, 3, 4 };
                List<int> Lane_Dist = DataProcessing.GetLaneDist(Lane.Text, dataPredicate, highSpeedData).ToList();
                string Lane_Dist_WriteString = string.Empty;
                try
                {
                    var fs = new FileStream("不同车道车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Lane_Dist[0]}";
                    for (int i = 1; i < Lane_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{Lane_Dist[i]}";
                    }
                    Lane_Dist_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


                //不同区间车速车数量分布
                string Speed_Dist_WriteString = string.Empty;
                List<int> Speed_Dist = DataProcessing.GetSpeedDist(Speed.Text, dataPredicate, highSpeedData).ToList();
                try
                {
                    var fs = new FileStream("不同车速区间车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Speed_Dist[0]}";
                    for (int i = 1; i < Speed_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{Speed_Dist[i]}";
                    }
                    Speed_Dist_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


                //var Gross_Load_Div = new int[] { 0, 10_000, 20_000, 30_000 };
                //不同区间车重车数量分布
                List<int> GrossLoad_Dist = DataProcessing.GetGrossLoadDist(GrossLoad.Text, dataPredicate, highSpeedData).ToList();

                string GrossLoad_Dist_WriteString = string.Empty;
                try    //结果写入txt（以逗号分隔）
                {
                    var fs = new FileStream("不同车重区间车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{GrossLoad_Dist[0]}";
                    for (int i = 1; i < GrossLoad_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{GrossLoad_Dist[i]}";
                    }
                    GrossLoad_Dist_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }



                //指定车道不同区间车重车数量分布

                int[] CriticalLane_Div = Array.ConvertAll(CriticalLane.Text.Split(','), s => int.Parse(s));

                string GrossLoadDistByLane_WriteString = string.Empty;
                for (int i = 0; i < CriticalLane_Div.Length; i++)
                {
                    List<int> GrossLoad_Dist_ByLane = DataProcessing.GetGrossLoadDistByLane(GrossLoad.Text, CriticalLane_Div[i], dataPredicate, highSpeedData).ToList();
                    try    //结果写入txt（以逗号分隔）
                    {
                        var fs = new FileStream($"车道{CriticalLane_Div[i]}不同车重区间车辆数.txt", FileMode.Create);
                        var sw = new StreamWriter(fs, Encoding.Default);
                        var writeString = $"{GrossLoad_Dist_ByLane[0]}";
                        for (int j = 1; j < GrossLoad_Dist_ByLane.Count; j++)
                        {
                            writeString = $"{writeString},{GrossLoad_Dist_ByLane[j]}";
                        }
                        GrossLoadDistByLane_WriteString = $"{GrossLoadDistByLane_WriteString }车道{CriticalLane_Div[i]}：{writeString}\n";

                        sw.Write(writeString);
                        sw.Close();
                        fs.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                string CustomWeightCount_WriteString = string.Empty;

                //指定车道大于指定车重车数量统计
                for (int i = 0; i < CriticalLane_Div.Length; i++)
                {
                    List<int> CustomWeightCount = DataProcessing.GetGrossLoadCountByLane(CustomWeight.Text, CriticalLane_Div[i], dataPredicate, highSpeedData).ToList();
                    try    //结果写入txt（以逗号分隔）
                    {
                        var fs = new FileStream($"车道{CriticalLane_Div[i]}大于指定车重车数量统计.txt", FileMode.Create);
                        var sw = new StreamWriter(fs, Encoding.Default);
                        var writeString = $"{CustomWeightCount[0]}({(decimal)CustomWeightCount[0] / Lane_Dist[i]:P})";
                        for (int j = 1; j < CustomWeightCount.Count; j++)
                        {
                            writeString = $"{writeString},{CustomWeightCount[j]}({(decimal)CustomWeightCount[j] / Lane_Dist[i]:P})";
                        }
                        CustomWeightCount_WriteString = $"{CustomWeightCount_WriteString }车道{CriticalLane_Div[i]}：{writeString}\n";
                        sw.Write(writeString);
                        sw.Close();
                        fs.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }



                //例子：
                //var Hour_Div = new int[] { 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22 };

                //不同区间小时车数量分布

                string Hour_Dist_WriteString = string.Empty;

                List<int> Hour_Dist = DataProcessing.GetHourDist(Hour.Text, dataPredicate, highSpeedData).ToList();
                try
                {
                    var fs = new FileStream("不同小时区间车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Hour_Dist[0]}";
                    for (int i = 1; i < Hour_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{Hour_Dist[i]}";
                    }
                    Hour_Dist_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


                //不同区间小时平均车速分布
                List<double?> HourSpeed_Dist = DataProcessing.GetHourAverageSpeedDist(Hour.Text, dataPredicate, highSpeedData).ToList();

                string HourSpeed_Dist_WriteString = string.Empty;

                try
                {
                    var fs = new FileStream("不同小时区间平均车速.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Math.Round(Convert.ToDecimal(HourSpeed_Dist[0] ?? 0.0), 1)}";
                    for (int i = 1; i < HourSpeed_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{Math.Round(Convert.ToDecimal(HourSpeed_Dist[i] ?? 0.0), 1)}";
                    }
                    HourSpeed_Dist_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                //不同区间小时大于指定重量的车数量
                var HourWeight_Dist = DataProcessing.GetCountsAboveCriticalWeightInDifferentHours(Hour.Text, Convert.ToInt32(CriticalWeight.Text), dataPredicate, highSpeedData).ToList();

                string HourWeight_Dist_WriteString = string.Empty;

                try
                {
                    var fs = new FileStream($"不同区间小时大于{CriticalWeight.Text}kg车数量.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{HourWeight_Dist[0]}";
                    for (int i = 1; i < HourWeight_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{HourWeight_Dist[i]}";
                    }
                    HourWeight_Dist_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


                //大于指定重量车辆数统计
                List<int> GrossLoadCount = DataProcessing.GetGrossLoadCount(CustomWeight.Text, dataPredicate, highSpeedData).ToList();
                string GrossLoadCount_WriteString = string.Empty;
                try
                {
                    var fs = new FileStream("大于指定车重车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{GrossLoadCount[0]}";
                    for (int i = 1; i < GrossLoadCount.Count; i++)
                    {
                        writeString = $"{writeString},{GrossLoadCount[i]}";
                    }
                    GrossLoadCount_WriteString = writeString;
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                SummaryString = $"不同车道车辆数：{Lane_Dist_WriteString}\n不同车速区间车辆数：{Speed_Dist_WriteString}\n不同车重区间车辆数：{GrossLoad_Dist_WriteString}\n指定车道不同区间车重车数量分布：{GrossLoadDistByLane_WriteString}\n指定车道大于指定车重车数量统计：{CustomWeightCount_WriteString}\n不同小时区间车辆数：{Hour_Dist_WriteString}\n不同区间小时平均车速分布：{HourSpeed_Dist_WriteString}\n不同区间小时大于指定重量的车数量：{HourWeight_Dist_WriteString}\n大于指定重量车辆数统计：{GrossLoadCount_WriteString}";

                try
                {
                    var fs = new FileStream("统计汇总.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    sw.Write(SummaryString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                ExportToExcelForPythonPlot(dataPredicate, highSpeedData, Lane_Dist_WriteString, Speed_Dist_WriteString, GrossLoad_Dist_WriteString, CriticalLane_Div, Hour_Dist_WriteString);

                ExportToDocx(startDateTime, finishDateTime, dataPredicate, highSpeedData, Lane_Dist, CriticalLane_Div);

                _ = MessageBox.Show("运行完成！");
            }
        }

        //导出结果到Excel用于后期python作图
        private void ExportToExcelForPythonPlot(Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData, string Lane_Dist_WriteString, string Speed_Dist_WriteString, string GrossLoad_Dist_WriteString, int[] CriticalLane_Div, string Hour_Dist_WriteString)
        {
            var WIMToPythonPlotFileName = "动态称重.xlsx";

            FileInfo file = new FileInfo(WIMToPythonPlotFileName);

            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(WIMToPythonPlotFileName);
            }


            var sheetName = "Sheet1";
            try
            {
                using (var package = new ExcelPackage(file))
                {
                    //var worksheet = package.Workbook.Worksheets.Add(sheetName);    //新建
                    //var worksheet = package.Workbook.Worksheets[sheetName];    //已有
                    var worksheet = package.Workbook.Worksheets.Add(sheetName);

                    int currRow = 1;

                    worksheet.Cells[1, 1].Value = "序号";
                    worksheet.Cells[1, 2].Value = "文件名";
                    worksheet.Cells[1, 3].Value = "数值";
                    worksheet.Cells[1, 4].Value = "数值类型";
                    worksheet.Cells[1, 5].Value = "x轴标签";
                    worksheet.Cells[1, 6].Value = "y轴标签";
                    worksheet.Cells[1, 6].Value = "x轴标签标注占比";


                    //不同车重区间车辆数
                    worksheet.Cells[2, 2].Value = "不同车重区间车辆数";
                    worksheet.Cells[2, 3].Value = GrossLoad_Dist_WriteString;
                    worksheet.Cells[2, 4].Value = "int";


                    //警告：仅支持整数
                    //new int[] { 0, 10_000, 20_000, 30_000 };
                    string tempStr = string.Empty;

                    string Gross_Load_Div_XlabelString = string.Empty;
                    int[] Gross_Load_Div = Array.ConvertAll(GrossLoad.Text.Split(','), s => int.Parse(s));

                    for (int i = 0; i < Gross_Load_Div.Length; i++)
                    {
                        if (i == 0)
                        {
                            Gross_Load_Div_XlabelString = $"0～{Gross_Load_Div[1] / 1000}t";
                        }
                        else if (i != Gross_Load_Div.Length - 1)
                        {
                            Gross_Load_Div_XlabelString = $"{Gross_Load_Div_XlabelString},{Gross_Load_Div[i] / 1000}～{Gross_Load_Div[i + 1] / 1000}t";
                        }
                        else
                        {
                            Gross_Load_Div_XlabelString = $"{Gross_Load_Div_XlabelString},{Gross_Load_Div[i] / 1000}t以上";
                        }

                    }

                    worksheet.Cells[2, 5].Value = Gross_Load_Div_XlabelString;
                    worksheet.Cells[2, 6].Value = "数量";
                    worksheet.Cells[2, 7].Value = "是";

                    //不同车道车辆数
                    worksheet.Cells[3, 2].Value = "不同车道车辆数";
                    worksheet.Cells[3, 3].Value = Lane_Dist_WriteString;
                    worksheet.Cells[3, 4].Value = "int";

                    tempStr = string.Empty;
                    int[] Lane_Div = Array.ConvertAll(Lane.Text.Split(','), s => int.Parse(s));
                    for (int i = 0; i < Lane_Div.Length; i++)
                    {
                        if (i == 0)
                        {
                            tempStr = $"车道{Lane_Div[0]}";
                        }
                        else
                        {
                            tempStr = $"{tempStr},车道{Lane_Div[i]}";
                        }

                    }

                    string Lane_Dist_Xlabeltring = tempStr;
                    worksheet.Cells[3, 5].Value = Lane_Dist_Xlabeltring;
                    worksheet.Cells[3, 6].Value = "数量";
                    worksheet.Cells[3, 7].Value = "是";

                    //不同车速区间车辆数
                    worksheet.Cells[4, 2].Value = "不同车速区间车辆数";
                    worksheet.Cells[4, 3].Value = Speed_Dist_WriteString;
                    worksheet.Cells[4, 4].Value = "int";

                    tempStr = string.Empty;
                    int[] Speed_Div = Array.ConvertAll(Speed.Text.Split(','), s => int.Parse(s));
                    for (int i = 0; i < Speed_Div.Length; i++)
                    {
                        if (i == 0)
                        {
                            tempStr = $"0～{Speed_Div[1]}km/h";
                        }
                        else if (i != Speed_Div.Length - 1)
                        {
                            tempStr = $"{tempStr},{Speed_Div[i]}～{Speed_Div[i + 1] }km/h";
                        }
                        else
                        {
                            tempStr = $"{tempStr},{Speed_Div[i]}km/h以上";
                        }

                    }

                    worksheet.Cells[4, 5].Value = tempStr;
                    worksheet.Cells[4, 6].Value = "数量";
                    worksheet.Cells[4, 7].Value = "是";

                    //不同小时区间车辆数
                    worksheet.Cells[5, 2].Value = "不同小时区间车辆数";
                    worksheet.Cells[5, 3].Value = Hour_Dist_WriteString;
                    worksheet.Cells[5, 4].Value = "int";

                    tempStr = string.Empty;
                    int[] Hour_Div = Array.ConvertAll(Hour.Text.Split(','), s => int.Parse(s));
                    for (int i = 0; i < Hour_Div.Length; i++)
                    {
                        if (i == 0)
                        {
                            tempStr = $"0～{Hour_Div[1]}h";
                        }
                        else if (i != Hour_Div.Length - 1)
                        {
                            tempStr = $"{tempStr},{Hour_Div[i]}～{Hour_Div[i + 1] }";
                        }
                        else
                        {
                            tempStr = $"{tempStr},{Hour_Div[i]}～24";
                        }
                    }

                    worksheet.Cells[5, 5].Value = tempStr;
                    worksheet.Cells[5, 6].Value = "数量";
                    worksheet.Cells[5, 7].Value = "否";

                    currRow = 5;

                    //车道x不同车重区间车辆数
                    for (int i = 0; i < CriticalLane_Div.Length; i++)
                    {
                        worksheet.Cells[currRow, 2].Value = $"车道{CriticalLane_Div[i]}不同车重区间车辆数";

                        List<int> GrossLoad_Dist_ByLane = DataProcessing.GetGrossLoadDistByLane(GrossLoad.Text, CriticalLane_Div[i], dataPredicate, highSpeedData).ToList();


                        worksheet.Cells[currRow, 3].Value = string.Join(",", GrossLoad_Dist_ByLane);
                        worksheet.Cells[currRow, 4].Value = "int";

                        worksheet.Cells[currRow, 5].Value = Gross_Load_Div_XlabelString;
                        worksheet.Cells[currRow, 6].Value = "数量";
                        worksheet.Cells[currRow, 7].Value = "否";
                        currRow++;
                    }


                    package.Save();
                }

            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }

        //导出结果到docx
        private void ExportToDocx(DateTime startDateTime, DateTime finishDateTime, Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData, List<int> Lane_Dist, int[] CriticalLane_Div)
        {
            string templateFile = $"{App.TemplateFolder}\\大樟桥监测月报模板.docx";
            string outputFile = $"{App.OutputFolder}\\{DateTime.Now:yyyyMMdd}自动生成的大樟桥监测月报.docx";

            var ShortFinishDateTime = finishDateTime;
            var k = highSpeedData.Where(dataPredicate);

            int TotalVehicleCount = highSpeedData.Where(dataPredicate).Count();
            int Lane1VehicleCount = highSpeedData.Where(dataPredicate).Where(x => x.Lane_Id == 1).Count();
            int Lane2VehicleCount = highSpeedData.Where(dataPredicate).Where(x => x.Lane_Id == 2).Count();
            int DailyVehicleCount = Convert.ToInt32(Math.Round(TotalVehicleCount * 1.0m / (finishDateTime - startDateTime).Days));

            List<int> CustomWeightCountx = DataProcessing.GetGrossLoadCountByLane(CustomWeight.Text, CriticalLane_Div[0], dataPredicate, highSpeedData).ToList();
            int Lane1Vehicle30Count = CustomWeightCountx[0];
            decimal Lane1Vehicle30Proportion = (decimal)CustomWeightCountx[0] / Lane_Dist[0];
            int Lane1Vehicle55Count = CustomWeightCountx[1];
            decimal Lane1Vehicle55Proportion = (decimal)CustomWeightCountx[1] / Lane_Dist[0];

            CustomWeightCountx = DataProcessing.GetGrossLoadCountByLane(CustomWeight.Text, CriticalLane_Div[1], dataPredicate, highSpeedData).ToList();

            int Lane2Vehicle30Count = CustomWeightCountx[0];
            decimal Lane2Vehicle30Proportion = (decimal)CustomWeightCountx[0] / Lane_Dist[1];
            int Lane2Vehicle55Count = CustomWeightCountx[1];
            decimal Lane2Vehicle55Proportion = (decimal)CustomWeightCountx[1] / Lane_Dist[1];

            try
            {
                var doc = new Document(templateFile);

                string[] MyDocumentVariables = new string[] { nameof(StartDateTime), nameof(FinishDateTime), nameof(ShortFinishDateTime), nameof(TotalVehicleCount), nameof(Lane1VehicleCount), nameof(Lane2VehicleCount), nameof(DailyVehicleCount), nameof(Lane1Vehicle30Count), nameof(Lane1Vehicle30Proportion), nameof(Lane1Vehicle55Count), nameof(Lane1Vehicle55Proportion), nameof(Lane2Vehicle30Count), nameof(Lane2Vehicle30Proportion), nameof(Lane2Vehicle55Count), nameof(Lane2Vehicle55Proportion) };//文档中包含的所有“文档变量”，方便遍历
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   //更新文档变量
                try
                {
                    var variables = doc.Variables;
                    variables[nameof(StartDateTime)] = startDateTime.ToString("yyyy年MM月dd日");
                    variables[nameof(FinishDateTime)] = finishDateTime.AddDays(-1).ToString("yyyy年MM月dd日");
                    variables[nameof(ShortFinishDateTime)] = finishDateTime.AddDays(-1).ToString("MM月dd日");
                    variables[nameof(TotalVehicleCount)] = TotalVehicleCount.ToString();
                    variables[nameof(Lane1VehicleCount)] = Lane1VehicleCount.ToString();
                    variables[nameof(Lane2VehicleCount)] = Lane2VehicleCount.ToString();
                    variables[nameof(DailyVehicleCount)] = DailyVehicleCount.ToString();
                    variables[nameof(Lane1Vehicle30Count)] = Lane1Vehicle30Count.ToString();
                    variables[nameof(Lane1Vehicle30Proportion)] = Lane1Vehicle30Proportion.ToString("P");
                    variables[nameof(Lane1Vehicle55Count)] = Lane1Vehicle55Count.ToString();
                    variables[nameof(Lane1Vehicle55Proportion)] = Lane1Vehicle55Proportion.ToString("P");
                    variables[nameof(Lane2Vehicle30Count)] = Lane2Vehicle30Count.ToString();
                    variables[nameof(Lane2Vehicle30Proportion)] = Lane2Vehicle30Proportion.ToString("P");
                    variables[nameof(Lane2Vehicle55Count)] = Lane2Vehicle55Count.ToString();
                    variables[nameof(Lane2Vehicle55Proportion)] = Lane2Vehicle55Proportion.ToString("P");
                }
                catch (Exception ex)
                {

                    Debug.Print($"文档变量更新失败。信息{ex.Message}");
                }

                doc.UpdateFields();

                foreach (var v in doc.Range.Fields)
                {
                    FieldDocVariable v1 = v as FieldDocVariable;
                    if (v1 != null)
                    {
                        if (MyDocumentVariables.Contains(v1.VariableName))
                        {
                            v1.Unlink();
                        }
                    }

                }

                var builder = new DocumentBuilder(doc);

                try
                {
                    var topGrossLoadtable1 = doc.GetChildNodes(NodeType.Table, true)[6] as Aspose.Words.Tables.Table;
                    var topGrossLoadtable2 = doc.GetChildNodes(NodeType.Table, true)[7] as Aspose.Words.Tables.Table;    //续上一张表
                    List<HighSpeedData> data = highSpeedData.Where(dataPredicate).OrderByDescending(x => x.Gross_Load).Take(Convert.ToInt32(10)).ToList();

                    for (int i = 0; i < 10; i++)
                    {

                        builder.MoveTo(topGrossLoadtable1.Rows[i + 1].Cells[1].FirstParagraph);
                        builder.Write($"{data[i].Lane_Id}");
                        builder.MoveTo(topGrossLoadtable1.Rows[i + 1].Cells[2].FirstParagraph);
                        builder.Write($"{data[i].HSData_DT}");
                        builder.MoveTo(topGrossLoadtable1.Rows[i + 1].Cells[3].FirstParagraph);
                        builder.Write($"{data[i].Axle_Num}");
                        builder.MoveTo(topGrossLoadtable1.Rows[i + 1].Cells[4].FirstParagraph);
                        builder.Write($"{data[i].Gross_Load}");
                        builder.MoveTo(topGrossLoadtable1.Rows[i + 1].Cells[5].FirstParagraph);
                        builder.Write($"{data[i].Speed}");

                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[1].FirstParagraph);
                        builder.Write($"{data[i].LWheel_1_W + data[i].RWheel_1_W}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[2].FirstParagraph);
                        builder.Write($"{data[i].LWheel_2_W + data[i].RWheel_2_W}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[3].FirstParagraph);
                        builder.Write($"{data[i].LWheel_3_W + data[i].RWheel_3_W}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[4].FirstParagraph);
                        builder.Write($"{data[i].LWheel_4_W + data[i].RWheel_4_W}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[5].FirstParagraph);
                        builder.Write($"{data[i].LWheel_5_W + data[i].RWheel_5_W}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[6].FirstParagraph);
                        builder.Write($"{data[i].LWheel_6_W + data[i].RWheel_6_W}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[7].FirstParagraph);
                        builder.Write($"{Math.Round(Convert.ToDecimal(data[i].AxleDis1 ?? 0.0) / 1000, 2)}"); ;
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[8].FirstParagraph);
                        builder.Write($"{Math.Round(Convert.ToDecimal(data[i].AxleDis2 ?? 0.0) / 1000, 2)}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[9].FirstParagraph);
                        builder.Write($"{Math.Round(Convert.ToDecimal(data[i].AxleDis3 ?? 0.0) / 1000, 2)}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[10].FirstParagraph);
                        builder.Write($"{Math.Round(Convert.ToDecimal(data[i].AxleDis4 ?? 0.0) / 1000, 2)}");
                        builder.MoveTo(topGrossLoadtable2.Rows[i + 1].Cells[11].FirstParagraph);
                        builder.Write($"{Math.Round(Convert.ToDecimal(data[i].AxleDis5 ?? 0.0) / 1000, 2)}");
                    }
                }
                catch (Exception ex)
                {

                    Debug.Print($"警告：前10最重车辆参数信息插入失败。信息{ex.Message}");
                }



                doc.UpdateFields();
                doc.Save(outputFile, SaveFormat.Docx);
                _ = MessageBox.Show("成功导出报告！");
            }
            catch (Exception ex)
            {

                Debug.Print($"警告：报告生成失败。信息{ex.Message}");
            }
        }

        private void ExportDailyTrafficVolumeAndExportTopGrossLoad_Click(object sender, RoutedEventArgs e)
        {
            ExportDailyTrafficVolume();
            ExportTopGrossLoad();
            MessageBox.Show("运行完成！");

        }

        private void ExportDailyTrafficVolume_Click(object sender, RoutedEventArgs e)
        {
            ExportDailyTrafficVolume();

            MessageBox.Show("运行完成！");

        }

        private void ExportDailyTrafficVolume()
        {
            var startDateTime = StartDateTime.SelectedDate ?? DateTime.Now.AddDays(-1);
            var finishDateTime = FinishDateTime.SelectedDate ?? DateTime.Now;
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDateTime && x.HSData_DT < finishDateTime;

            using (var db = new HighSpeed_PROCEntities())
            {

                var highSpeedData = (
                    from e1 in db.HS_Data_PROC
                    select new HighSpeedData
                    {
                        Acceleration = e1.Acceleration,
                        AxleGrp_Num = e1.AxleGrp_Num,
                        Axle_Num = e1.Axle_Num,
                        ExternInfo = e1.ExternInfo,
                        F7Code = e1.F7Code,
                        HSData_Id = e1.HSData_Id,
                        Veh_Length = e1.Veh_Length,
                        Veh_Type = e1.Veh_Type,
                        LWheel_1_W = e1.LWheel_1_W,
                        LWheel_2_W = e1.LWheel_2_W,
                        LWheel_3_W = e1.LWheel_3_W,
                        LWheel_4_W = e1.LWheel_4_W,
                        LWheel_5_W = e1.LWheel_5_W,
                        LWheel_6_W = e1.LWheel_6_W,
                        LWheel_7_W = e1.LWheel_7_W,
                        LWheel_8_W = e1.LWheel_8_W,
                        Lane_Id = e1.Lane_Id,
                        Oper_Direc = e1.Oper_Direc,
                        Speed = e1.Speed,
                        OverLoad_Sign = e1.OverLoad_Sign,
                        RWheel_1_W = e1.RWheel_1_W,
                        RWheel_2_W = e1.RWheel_2_W,
                        RWheel_3_W = e1.RWheel_3_W,
                        RWheel_4_W = e1.RWheel_4_W,
                        RWheel_5_W = e1.RWheel_5_W,
                        RWheel_6_W = e1.RWheel_6_W,
                        RWheel_7_W = e1.RWheel_7_W,
                        RWheel_8_W = e1.RWheel_8_W,
                        Violation_Id = e1.Violation_Id,
                        AxleDis1 = e1.AxleDis1,
                        AxleDis2 = e1.AxleDis2,
                        AxleDis3 = e1.AxleDis3,
                        AxleDis4 = e1.AxleDis4,
                        AxleDis5 = e1.AxleDis5,
                        AxleDis6 = e1.AxleDis6,
                        AxleDis7 = e1.AxleDis7,
                        HSData_DT = e1.HSData_DT,
                        Gross_Load = e1.Gross_Load,
                        License_Plate=e1.License_Plate
                    }
                    );
                IEnumerable<DailyTraffic> dailyTrafficData = DataProcessing.GetDailyTraffic(Lane.Text, highSpeedData, startDateTime, finishDateTime);

                //每日交通流量信息数据导入excel
                var temp = ExcelHelper.ExportDailyTrafficVolume(dailyTrafficData.ToList());
            }
        }

        //重量前10的车辆数据导入excel
        private void ExportTopGrossLoad_Click(object sender, RoutedEventArgs e)
        {
            ExportTopGrossLoad();

            MessageBox.Show("运行完成！");

        }

        private void ExportTopGrossLoad()
        {
            var startDateTime = StartDateTime.SelectedDate ?? DateTime.Now.AddDays(-1);
            var finishDateTime = FinishDateTime.SelectedDate ?? DateTime.Now;
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDateTime && x.HSData_DT < finishDateTime;

            using (var db = new HighSpeed_PROCEntities())
            {

                var highSpeedData = (
                    from e1 in db.HS_Data_PROC
                    select new HighSpeedData
                    {
                        Acceleration = e1.Acceleration,
                        AxleGrp_Num = e1.AxleGrp_Num,
                        Axle_Num = e1.Axle_Num,
                        ExternInfo = e1.ExternInfo,
                        F7Code = e1.F7Code,
                        HSData_Id = e1.HSData_Id,
                        Veh_Length = e1.Veh_Length,
                        Veh_Type = e1.Veh_Type,
                        LWheel_1_W = e1.LWheel_1_W,
                        LWheel_2_W = e1.LWheel_2_W,
                        LWheel_3_W = e1.LWheel_3_W,
                        LWheel_4_W = e1.LWheel_4_W,
                        LWheel_5_W = e1.LWheel_5_W,
                        LWheel_6_W = e1.LWheel_6_W,
                        LWheel_7_W = e1.LWheel_7_W,
                        LWheel_8_W = e1.LWheel_8_W,
                        Lane_Id = e1.Lane_Id,
                        Oper_Direc = e1.Oper_Direc,
                        Speed = e1.Speed,
                        OverLoad_Sign = e1.OverLoad_Sign,
                        RWheel_1_W = e1.RWheel_1_W,
                        RWheel_2_W = e1.RWheel_2_W,
                        RWheel_3_W = e1.RWheel_3_W,
                        RWheel_4_W = e1.RWheel_4_W,
                        RWheel_5_W = e1.RWheel_5_W,
                        RWheel_6_W = e1.RWheel_6_W,
                        RWheel_7_W = e1.RWheel_7_W,
                        RWheel_8_W = e1.RWheel_8_W,
                        Violation_Id = e1.Violation_Id,
                        AxleDis1 = e1.AxleDis1,
                        AxleDis2 = e1.AxleDis2,
                        AxleDis3 = e1.AxleDis3,
                        AxleDis4 = e1.AxleDis4,
                        AxleDis5 = e1.AxleDis5,
                        AxleDis6 = e1.AxleDis6,
                        AxleDis7 = e1.AxleDis7,
                        HSData_DT = e1.HSData_DT,
                        Gross_Load = e1.Gross_Load,
                        License_Plate=e1.License_Plate
                        
                    }
                    );

                List<HighSpeedData> data = highSpeedData.Where(dataPredicate).OrderByDescending(x => x.Gross_Load).Take(Convert.ToInt32(CriticalCount.Text)).ToList();
                var temp = ExcelHelper.ExportTopGrossLoad(data);
            }
        }
    }

}
