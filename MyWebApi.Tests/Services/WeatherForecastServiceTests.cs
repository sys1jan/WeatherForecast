using Xunit;
using Moq;
using MyWebApi.Services;
using MyWebApi.Models;
using Google.Cloud.Firestore;
using System.Data.Common;

namespace MyWebApi.Tests.Services
{
    public class WeatherForecastServiceTests
    {
        private readonly Mock<IFirestoreDb> _dbMock;
        private readonly WeatherForecastService _service;
        private readonly Mock<Query> _queryMock;
        
        //private readonly CollectionReference _collectionMock;

        public WeatherForecastServiceTests()
        {
            _dbMock = new Mock<IFirestoreDb>();
            _queryMock = new Mock<Query>();
           // _collectionMock = 
            //_dbMock.Setup(db => db.Collection(It.IsAny<string>())).Returns(_collectionMock.Object);
            _service = new WeatherForecastService(_dbMock.Object);
        }

        [Fact]
        public async Task GetForecastsAsync_ReturnsExpectedResults()
        {
            // Arrange
            var mockDocument = new Mock<DocumentSnapshot>();
            mockDocument.Setup(d => d.Exists).Returns(true);
            mockDocument.Setup(d => d.ConvertTo<WeatherForecast>()).Returns(new WeatherForecast { Date = "2021-01-01 00:00:00", Name = "Test", TemperatureC = 0, TemperatureF = 32, Summary = new List<string> { "Test" }, Description = new List<string> { "Test" } });

            var mockSnapshot = new Mock<QuerySnapshot>();
            mockSnapshot.Setup(s => s.Documents).Returns(new List<DocumentSnapshot> { mockDocument.Object });

            _queryMock.Setup(q => q.GetSnapshotAsync(It.IsAny<CancellationToken>())).ReturnsAsync(mockSnapshot.Object);

            // Act
            var result = await _service.GetForecastsAsync();

            // Assert
            // TODO: Add your assertions here
            Assert.Single(result);
            Assert.Equal("Test", result[0].Name);
            Assert.Equal(0, result[0].TemperatureC);
            Assert.Equal(32, result[0].TemperatureF);
            Assert.Equal("Test", result[0].Summary[0]);
            Assert.Equal("Test", result[0].Description[0]);
            Assert.Equal("2021-01-01 00:00:00", result[0].Date);
            // TODO: Add more assertions
        }
        // TODO: Add more tests for InsertForecast, UpdateForecast, and GetForecast
    }
}