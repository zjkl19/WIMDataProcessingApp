using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace WIMDataProcessingApp
{
    public class ExcelHelper
    {
        /// <summary>
        /// 将每日交通流量数据导入excel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public static int ExportDailyTrafficVolume(List<DailyTraffic> data)
        {
            var file = new FileInfo("导出的交通流量数据.xlsx");    //文件要提前建立
            var sheetName = "sheet1";
            try
            {
                using (var package = new ExcelPackage(file))
                {
                    //var worksheet = package.Workbook.Worksheets.Add(sheetName);    //新建
                    var worksheet = package.Workbook.Worksheets[sheetName];    //已有
                    worksheet.Cells[2, 1].Value = "上行车数";
                    worksheet.Cells[3, 1].Value = "下行车数";
                    worksheet.Cells[4, 1].Value = "车辆总数";
                    worksheet.Cells[5, 1].Value = "上行方向车辆总数"; worksheet.Cells[5, 2].Value = data.Sum(x => x.UpStreamCount);
                    worksheet.Cells[6, 1].Value = "上行方向车辆总数"; worksheet.Cells[6, 2].Value = data.Sum(x => x.DownStreamCount);
                    worksheet.Cells[7, 1].Value = "日均车辆数"; worksheet.Cells[7, 2].Value = Math.Round(data.Sum(x => x.TotalStreamCount) * 1.0m / data.Count);
                    for (int i = 0; i < data.Count; i++)
                    {
                        worksheet.Cells[1, i + 2].Value = data[i].Date;
                        worksheet.Cells[2, i + 2].Value = data[i].UpStreamCount;
                        worksheet.Cells[3, i + 2].Value = data[i].DownStreamCount;
                        worksheet.Cells[4, i + 2].Value = data[i].TotalStreamCount;
                    }
                    package.Save();
                }
                return 1;
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
                return 0;
            }
        }



    }
}
