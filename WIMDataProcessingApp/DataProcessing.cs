using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace WIMDataProcessingApp
{
    public class DataProcessing
    {
        public static IEnumerable<DailyTraffic> GetDailyTraffic(IQueryable<HighSpeedData> table, DateTime StartDataTime, DateTime FinishDataTime)
        {
            DailyTraffic dailyTraffic;
            var currTime = StartDataTime;
            //var cc = table.Where(x => EntityFunctions.TruncateTime(x.HSData_DT)>= EntityFunctions.TruncateTime(StartDataTime)).Count();
            while (currTime.AddDays(1) <= FinishDataTime)
            {
                var k1 = currTime;
                var k2 = currTime.AddDays(1);

                var cc = table.Where(x => x.HSData_DT >= k1 && x.HSData_DT < k2).Count();
                dailyTraffic = new DailyTraffic
                {
                    Date = currTime.ToString("m")
                    ,
                    UpStreamCount = table.Where(x => x.HSData_DT >= k1 && x.HSData_DT < k2 && (x.Lane_Id == 1 || x.Lane_Id == 2)).Count()
                    ,
                    DownStreamCount = table.Where(x => x.HSData_DT >= k1 && x.HSData_DT < k2 && (x.Lane_Id == 3 || x.Lane_Id == 4)).Count()
                    ,
                    TotalStreamCount = table.Where(x => x.HSData_DT >= k1 && x.HSData_DT < k2).Count()

                };
                currTime = k2;
                yield return dailyTraffic;
            }


        }

        //不同车道车数量分布
        public static IEnumerable<int> GetLaneDist(string laneText,Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData)
        {
            int temp;
            int[] Lane_Div = Array.ConvertAll(laneText.Split(','), s => int.Parse(s));
            for (int i = 0; i < Lane_Div.Length; i++)
            {
                temp = Lane_Div[i];
                yield return highSpeedData.Where(x => x.Lane_Id == temp).Where(dataPredicate).Count();
            }
        }
        //不同区间车速车数量分布
        public static IEnumerable<int> GetSpeedDist(string speedText, Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData)
        {
            int t1,t2;    //临时变量
            int[] Speed_Div = Array.ConvertAll(speedText.Split(','), s => int.Parse(s));
            //var Speed_Dist = new List<int>();
            
            for (int i = 0; i < Speed_Div.Length; i++)
            {
                t1 = Speed_Div[i];
                if (i != Speed_Div.Length - 1)
                {
                    t2 = Speed_Div[i + 1];
                    yield return highSpeedData.Where(x => x.Speed >= t1 && x.Speed < t2).Where(dataPredicate).Count();
                }
                else
                {
                    yield return highSpeedData.Where(x => x.Speed >= t1).Where(dataPredicate).Count();
                }
            }
        }
        //不同区间车重车数量分布
        public static IEnumerable<int> GetGrossLoadDist(string grossLoadText, Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData)
        {
            int t1, t2;    //临时变量
            int[] Gross_Load_Div = Array.ConvertAll(grossLoadText.Split(','), s => int.Parse(s));
            for (int i = 0; i < Gross_Load_Div.Length; i++)
            {
                t1 = Gross_Load_Div[i];
                if (i != Gross_Load_Div.Length - 1)
                {
                    t2 = Gross_Load_Div[i + 1];
                    yield return highSpeedData.Where(x => x.Gross_Load >= t1 && x.Gross_Load < t2).Where(dataPredicate).Count();
                }
                else
                {
                    yield return highSpeedData.Where(x => x.Gross_Load >= t1).Where(dataPredicate).Count();
                }
            }
        }
        //不同区间小时车数量分布
        public static IEnumerable<int> GetHourDist(string hourText, Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData)
        {
            int t1, t2;    //临时变量
            int[] Hour_Div = Array.ConvertAll(hourText.Split(','), s => int.Parse(s));
            var Hour_Dist = new List<int>();

            //不同区间小时分布
            for (int i = 0; i < Hour_Div.Length; i++)
            {
                t1 = Hour_Div[i];
                if (i != Hour_Div.Length - 1)
                {
                    t2 = Hour_Div[i + 1];
                    //TODO:Convert.ToDateTime(x.HSData_DT)中x.HSData_DT为空时会抛出异常
                    yield return highSpeedData.Where(x => x.HSData_DT.Value.Hour >= t1 && x.HSData_DT.Value.Hour < t2).Where(dataPredicate).Count();
                }
                else
                {
                    yield return highSpeedData.Where(x => x.HSData_DT.Value.Hour >= t1).Where(dataPredicate).Count();
                }
            }
        }
        //不同区间小时平均车速分布
        public static IEnumerable<double?> GetHourAverageSpeedDist(string hourText, Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData)
        {
            int t1, t2;    //临时变量
            int[] Hour_Div = Array.ConvertAll(hourText.Split(','), s => int.Parse(s));

            for (int i = 0; i < Hour_Div.Length; i++)
            {
                t1 = Hour_Div[i];
                if (i != Hour_Div.Length - 1)
                {
                    t2 = Hour_Div[i + 1];
                    //TODO:Convert.ToDateTime(x.HSData_DT)中x.HSData_DT为空时会抛出异常
                    yield return highSpeedData.Where(x => x.HSData_DT.Value.Hour >= t1 && x.HSData_DT.Value.Hour < t2).Where(dataPredicate).Average(x => x.Speed);
                }
                else
                {
                    yield return highSpeedData.Where(x => x.HSData_DT.Value.Hour >= t1).Where(dataPredicate).Average(x => x.Speed);
                }

            }
        }

        public static IEnumerable<int> GetCountsAboveCriticalWeightInDifferentHours(string hourText, int criticalWeight, Expression<Func<HighSpeedData, bool>> dataPredicate, IQueryable<HighSpeedData> highSpeedData)
        {
            int t1, t2;    //临时变量
            int[] Hour_Div = Array.ConvertAll(hourText.Split(','), s => int.Parse(s));

            for (int i = 0; i < Hour_Div.Length; i++)
            {
                t1 = Hour_Div[i];
                if (i != Hour_Div.Length - 1)
                {
                    t2 = Hour_Div[i + 1];
                    //TODO:Convert.ToDateTime(x.HSData_DT)中x.HSData_DT为空时会抛出异常
                    yield return highSpeedData.Where(x => x.HSData_DT.Value.Hour >= t1 && x.HSData_DT.Value.Hour < t2).Where(dataPredicate).Where(x => x.Gross_Load > criticalWeight).Count();
                }
                else
                {
                    yield return highSpeedData.Where(x => x.HSData_DT.Value.Hour >= t1).Where(dataPredicate).Where(x => x.Gross_Load > criticalWeight).Count();
                }

            }
        }
    }
}
