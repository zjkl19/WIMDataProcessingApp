using System;
using System.Collections.Generic;
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

            var startDataTime = StartDataTime.SelectedDate ?? DateTime.Now.AddDays(-1);
            var finishDataTime = FinishDataTime.SelectedDate ?? DateTime.Now;
            //var StartDataTime = new DateTime(2020, 10, 1, 0, 0, 0);
            //var FinishDataTime = new DateTime(2020, 11, 1, 0, 0, 0);
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDataTime && x.HSData_DT < finishDataTime;

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
                        Gross_Load = e1.Gross_Load
                    }
                    );
                #endregion;

                //不同车道车数量分布
                //var Lane_Div = new int[] { 1, 2, 3, 4 };
                List<int> Lane_Dist = DataProcessing.GetLaneDist(Lane.Text,dataPredicate, highSpeedData).ToList();
                try
                {
                    var fs = new FileStream("不同车道车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Lane_Dist[0]}";
                    for (int i = 1; i < Lane_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{Lane_Dist[i]}";
                    }
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


                //不同区间车速车数量分布
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
                try    //结果写入txt（以逗号分隔）
                {
                    var fs = new FileStream("不同车重区间车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{GrossLoad_Dist[0]}";
                    for (int i = 1; i < GrossLoad_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{GrossLoad_Dist[i]}";
                    }
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                //指定车道不同区间车重车数量分布

                int[] GrossLoad_Dist_ByLane_Div = Array.ConvertAll(CriticalLane.Text.Split(','), s => int.Parse(s));

                for (int i = 0; i < GrossLoad_Dist_ByLane_Div.Length; i++)
                {
                    List<int> GrossLoad_Dist_ByLane = DataProcessing.GetGrossLoadDistByLane(GrossLoad.Text, GrossLoad_Dist_ByLane_Div[i], dataPredicate, highSpeedData).ToList();
                    try    //结果写入txt（以逗号分隔）
                    {
                        var fs = new FileStream($"车道{GrossLoad_Dist_ByLane_Div[i]}不同车重区间车辆数.txt", FileMode.Create);
                        var sw = new StreamWriter(fs, Encoding.Default);
                        var writeString = $"{GrossLoad_Dist_ByLane[0]}";
                        for (int j = 1; j < GrossLoad_Dist_ByLane.Count; j++)
                        {
                            writeString = $"{writeString},{GrossLoad_Dist_ByLane[j]}";
                        }
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

                try
                {
                    var fs = new FileStream("不同小时区间平均车速.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Math.Round(Convert.ToDecimal(HourSpeed_Dist[0] ?? 0.0), 1)}";
                    for (int i = 1; i < HourSpeed_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{Math.Round(Convert.ToDecimal(HourSpeed_Dist[i] ?? 0.0), 1)}";
                    }
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                //不同区间小时大于指定重量的车数量
                var HourWeight_Dist = DataProcessing.GetCountsAboveCriticalWeightInDifferentHours(Hour.Text, Convert.ToInt32(CriticalWeight.Text),dataPredicate, highSpeedData).ToList();

                try
                {
                    var fs = new FileStream($"不同区间小时大于{CriticalWeight.Text}kg车数量.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{HourWeight_Dist[0]}";
                    for (int i = 1; i < HourWeight_Dist.Count; i++)
                    {
                        writeString = $"{writeString},{HourWeight_Dist[i]}";
                    }
                    sw.Write(writeString);
                    sw.Close();
                    fs.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                MessageBox.Show("运行完成！");
            }
        }


        private void ExportDailyTrafficVolume_Click(object sender, RoutedEventArgs e)
        {
            var startDataTime = StartDataTime.SelectedDate ?? DateTime.Now.AddDays(-1);
            var finishDataTime = FinishDataTime.SelectedDate ?? DateTime.Now;
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDataTime && x.HSData_DT < finishDataTime;

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
                        Gross_Load = e1.Gross_Load
                    }
                    );
                IEnumerable<DailyTraffic> dailyTrafficData = DataProcessing.GetDailyTraffic(Lane.Text,highSpeedData, startDataTime, finishDataTime);

                //每日交通流量信息数据导入excel
                var temp = ExcelHelper.ExportDailyTrafficVolume(dailyTrafficData.ToList());
            }

            MessageBox.Show("运行完成！");

        }
        //重量前10的车辆数据导入excel
        private void ExportTopGrossLoad_Click(object sender, RoutedEventArgs e)
        {
            var startDataTime = StartDataTime.SelectedDate ?? DateTime.Now.AddDays(-1);
            var finishDataTime = FinishDataTime.SelectedDate ?? DateTime.Now;
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDataTime && x.HSData_DT < finishDataTime;

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
                        Gross_Load = e1.Gross_Load
                    }
                    );

                List<HighSpeedData> data = highSpeedData.Where(dataPredicate).OrderByDescending(x => x.Gross_Load).Take(Convert.ToInt32(CriticalCount.Text)).ToList();
                var temp = ExcelHelper.ExportTopGrossLoad(data);
            }
            
            MessageBox.Show("运行完成！");

        }
    }

}
