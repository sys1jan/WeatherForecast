using System;
using Google.Cloud.Firestore;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MyWebApi.Services;

public class Program
{
    public static void Main(string[] args)
    {
        //Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", "/Users/jeffneal/projects/dotnet_source/MyWebApi/Keys/cybersecurity-class-404013-6f8510532d40.json");
        string? credentialPath = Environment.GetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS");
        //Environment.SetEnvironmentVariable("GOOGLE_APPLICATION_CREDENTIALS", "/app/publish/Keys/cybersecurity-class-404013-6f8510532d40.json");

        var builder = WebApplication.CreateBuilder(args);
        // Add services to the container.
        builder.Services.AddScoped<IWeatherForecastService, WeatherForecastService>();
        builder.Services.AddSingleton(_ => FirestoreDb.Create("cybersecurity-class-404013"));

        builder.Services.AddControllers();
        
        // Configure CORS
        builder.Services.AddCors(options =>
        {
            options.AddPolicy("AllowSpecificOrigin",
                builder =>
                {
                    builder.WithOrigins("http://localhost:4200") // replace with your Angular app's URL
                           .AllowAnyHeader()
                           .AllowAnyMethod();
                });
        });
        // Configure logging
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole();
        //builder.Logging.AddDebug();

        // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
        builder.Services.AddEndpointsApiExplorer();
        builder.Services.AddSwaggerGen();
        

        var app = builder.Build();

        // Configure the HTTP request pipeline.
        //if (app.Environment.IsDevelopment())
        //{
            app.UseSwagger();
            app.UseSwaggerUI();
       // }

        app.UseHttpsRedirection();

        app.UseAuthorization();
        app.UseCors("AllowSpecificOrigin");

        app.MapControllers();

        app.Run();
    }
}