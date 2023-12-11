using Google.Cloud.Firestore; // Namespace for FirestoreDb
using MyWebApi.Models; 
using System;

namespace MyWebApi.Services 
{
    public class WeatherForecastService : IWeatherForecastService
    {
        private readonly FirestoreDb _db;
        private const string WeatherCollection = "Weather";

        public WeatherForecastService(FirestoreDb db)
        {
            _db = db;
        }

        public async Task<List<WeatherForecast>> GetForecastsAsync()
        {
            try
            {
                CollectionReference colRef = _db.Collection(WeatherCollection);
                QuerySnapshot snapshot = await colRef.GetSnapshotAsync();

                List<WeatherForecast> forecasts = new List<WeatherForecast>();

                foreach (DocumentSnapshot document in snapshot.Documents)
                {
                    if (document.Exists)
                    {
                        WeatherForecast weatherForecast = document.ConvertTo<WeatherForecast>();
                        forecasts.Add(weatherForecast);
                    }
                }

                return forecasts;
            }
            catch (Exception)
            {
                // Log the exception and rethrow, or handle it as appropriate for your application.
                throw;
            }
        }

        public async Task InsertForecast(WeatherForecast forecast)
        {
            try
            {
                CollectionReference colRef = _db.Collection(WeatherCollection);
                await colRef.AddAsync(forecast);
            }
            catch (Exception)
            {
                // Handle exception
                throw;
            }
        }

        public async Task UpdateForecast(string documentId, WeatherForecast forecast)
        {
            try
            {
                DocumentReference docRef = _db.Collection(WeatherCollection).Document(documentId);
                await docRef.SetAsync(forecast, SetOptions.Overwrite);
            }
            catch (Exception)
            {
                // Handle exception
                throw;
            }
        }

        public async Task<WeatherForecast?> GetForecast(string documentId)
        {
            try
            {
                DocumentReference docRef = _db.Collection(WeatherCollection).Document(documentId);
                DocumentSnapshot snapshot = await docRef.GetSnapshotAsync();
                if (snapshot.Exists)
                {
                    WeatherForecast weatherForecast = snapshot.ConvertTo<WeatherForecast>();
                    return weatherForecast;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                // Handle exception
                throw;
            }
        }

    }
}