using System;
using WIMDataProcessingApp;
using System.Linq;
using Xunit;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace WIMDataProcessingAppTestProject
{
    public class DataProcessingTest
    {
        [Fact]
        public void GetLaneDist_Test()
        {
            //Arrange
            List<HighSpeedData> highSpeedData = new List<HighSpeedData> {
            new HighSpeedData{
                Lane_Id=1,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
        },
            new HighSpeedData{
                Lane_Id=2,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
            },
            new HighSpeedData{
                Lane_Id=3,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
            },
            new HighSpeedData{
                Lane_Id=3,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
            },
            new HighSpeedData{
                Lane_Id=4,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
            },
            new HighSpeedData{
                Lane_Id=4,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
            },
            new HighSpeedData{
                Lane_Id=4,
                HSData_DT= new DateTime(2020, 10, 2, 0, 0, 0),
            },
            };
            var laneText = "1,2,3,4";
            var startDataTime = new DateTime(2020, 10, 1, 0, 0, 0);
            var finishDataTime = new DateTime(2020, 11, 1, 0, 0, 0);
            Expression<Func<HighSpeedData, bool>> dataPredicate = x => x.HSData_DT >= startDataTime && x.HSData_DT < finishDataTime;


            //Act
            var Lane_Dist = DataProcessing.GetLaneDist(laneText, dataPredicate, highSpeedData.AsQueryable()).ToList();

            //Assert
            Assert.Equal(2,Lane_Dist[2]);
            Assert.Equal(3, Lane_Dist[3]);
        }
    }
}
