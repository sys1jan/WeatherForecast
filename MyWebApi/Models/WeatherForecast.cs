using Google.Cloud.Firestore;

namespace MyWebApi.Models
{
    /// <summary>
    /// Represents a weather forecast.
    /// </summary>
    [FirestoreData]
    public class WeatherForecast
    {
        /// <summary>
        /// Document ID of the forecast.
        /// </summary>
        [FirestoreDocumentId]
        public string? Id { get; set; }

        /// <summary>
        /// Gets or sets the date of the forecast.
        /// </summary>
        [FirestoreProperty]
        public required string Date { get; set; }
        
        /// <summary>
        /// Gets or sets the name of the forecast.
        /// </summary>
        [FirestoreProperty]
        public required string Name { get; set; }
        
        /// <summary>
        /// Gets or sets the temperature in Celsius.
        /// </summary>
        [FirestoreProperty]
        public long TemperatureC { get; set; }
        
        /// <summary>
        /// Gets or sets the temperature in Fahrenheit.
        /// </summary>
        [FirestoreProperty]
        public long TemperatureF { get; set; }
        
        /// <summary>
        /// Gets or sets the summary of the forecast.
        /// </summary>
        [FirestoreProperty]
        public List<string> Summary { get; set; } = new List<string>();
        
        /// <summary>
        /// Gets or sets the detailed description of the forecast.
        /// </summary>
        [FirestoreProperty]
        public List<string> Description { get; set; } = new List<string>();
    }
}