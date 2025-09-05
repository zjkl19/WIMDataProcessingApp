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
            // Arrange
            var highSpeedData = new List<HighSpeedData> {
        new HighSpeedData{ Lane_Id=1, HSData_DT = new DateTime(2020,10,2,0,0,0) },
        new HighSpeedData{ Lane_Id=2, HSData_DT = new DateTime(2020,10,2,0,0,0) },
        new HighSpeedData{ Lane_Id=3, HSData_DT = new DateTime(2020,10,2,0,0,0) },
        new HighSpeedData{ Lane_Id=3, HSData_DT = new DateTime(2020,10,2,0,0,0) },
        new HighSpeedData{ Lane_Id=4, HSData_DT = new DateTime(2020,10,2,0,0,0) },
        new HighSpeedData{ Lane_Id=4, HSData_DT = new DateTime(2020,10,2,0,0,0) },
        new HighSpeedData{ Lane_Id=4, HSData_DT = new DateTime(2020,10,2,0,0,0) },
    };
            var laneText = "1,2,3,4";
            var start = new DateTime(2020, 10, 1, 0, 0, 0);
            var finish = new DateTime(2020, 11, 1, 0, 0, 0);
            Expression<Func<HighSpeedData, bool>> pred = x => x.HSData_DT >= start && x.HSData_DT < finish;

            // Act
            var laneDist = DataProcessing.GetLaneDist(laneText, pred, highSpeedData.AsQueryable()).ToList();

            // Assert
            Assert.Equal(1, laneDist[0]);
            Assert.Equal(1, laneDist[1]);
            Assert.Equal(2, laneDist[2]);
            Assert.Equal(3, laneDist[3]);
        }

    }
}
