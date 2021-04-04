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

        private void AutoReport_Click(object sender, RoutedEventArgs e)
        {
            int t1,t2;

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

                //不同车道分布
                //var Lane_Div = new int[] { 1, 2, 3, 4 };
                int[] Lane_Div = Array.ConvertAll(Lane.Text.Split(','), s => int.Parse(s));
                var Lane_Dist = new List<int>();
                for (int i = 0; i < Lane_Div.Length; i++)
                {
                    t1 = Lane_Div[i];
                    Lane_Dist.Add(highSpeedData.Where(x => x.Lane_Id == t1).Where(dataPredicate).Count());
                    Console.WriteLine(Lane_Dist[i]);
                }
                try
                {
                    var fs = new FileStream("不同车道车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Lane_Dist[0]}";
                    for (int i = 1; i < Lane_Div.Length; i++)
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


                int[] Speed_Div = Array.ConvertAll(Speed.Text.Split(','), s => int.Parse(s));
                var Speed_Dist = new List<int>();
                //不同区间车速分布
                for (int i = 0; i < Speed_Div.Length; i++)
                {
                    t1 = Speed_Div[i];
                    if (i != Speed_Div.Length - 1)
                    {
                        t2 = Speed_Div[i + 1];
                        Speed_Dist.Add(highSpeedData.Where(x => x.Speed >= t1 && x.Speed < t2).Where(dataPredicate).Count());
                    }
                    else
                    {
                        Speed_Dist.Add(highSpeedData.Where(x => x.Speed >= t1).Where(dataPredicate).Count());
                    }
                    Console.WriteLine(Speed_Dist[i]);
                }
                try
                {
                    var fs = new FileStream("不同车速区间车辆数.txt", FileMode.Create);
                    var sw = new StreamWriter(fs, Encoding.Default);
                    var writeString = $"{Speed_Dist[0]}";
                    for (int i = 1; i < Speed_Div.Length; i++)
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

                MessageBox.Show("运行完成！");
            }
        }

        private void OpenReport_Click(object sender, RoutedEventArgs e)
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
                IEnumerable<DailyTraffic> dailyTrafficData = DataProcessing.GetDailyTraffic(highSpeedData, startDataTime, finishDataTime);

                //每日交通流量信息数据导入excel
                var temp = ExcelHelper.ExportDailyTrafficVolume(dailyTrafficData.ToList());
            }

 

        }
    }

}
