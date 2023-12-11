using MyWebApi.Models;

public interface IWeatherForecastService
    {
        Task<WeatherForecast?> GetForecast(string id);
        Task<List<WeatherForecast>> GetForecastsAsync();
        Task InsertForecast(WeatherForecast forecast);
        Task UpdateForecast(string documentId, WeatherForecast forecast);
    }


