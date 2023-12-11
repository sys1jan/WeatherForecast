using Microsoft.AspNetCore.Mvc;
using MyWebApi.Services;
using MyWebApi.Models;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace MyWebApi.Controllers
{
    /// <summary>
    /// This controller is responsible for handling weather forecasts.
    /// </summary>
    [ApiController]
    [Route("[controller]")]
    [Produces("application/json")]
    public class WeatherForecastController : ControllerBase
    {
        private readonly IWeatherForecastService _weatherForecastService;
        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(IWeatherForecastService weatherForecastService, ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
            _weatherForecastService = weatherForecastService;
        }

        /// <summary>
        /// Get a list of weather forecasts.
        /// </summary>
        /// <returns>A list of weather forecasts.</returns>
        [HttpGet(Name = "GetWeatherForecast")]
        [ProducesResponseType(typeof(List<WeatherForecast>), 200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        public async Task<ActionResult> Get()
        {
            try
            {
                var stopwatch = Stopwatch.StartNew();

                List<WeatherForecast> forecasts = await _weatherForecastService.GetForecastsAsync();

                stopwatch.Stop();
                var elapsedMilliseconds = stopwatch.ElapsedMilliseconds;

                // Log the elapsed time for performance monitoring
                _logger.LogDebug($"Time taken for GetWeatherForecast: {elapsedMilliseconds}ms");
                _logger.LogDebug($"Number of weather forecasts returned: {forecasts.Count}");
                // Return the elapsedMilliseconds with the forecasts
                return Ok(new { elapsedMilliseconds, forecasts });
            }
            catch (Exception ex)
            {
                // Log the exception for debugging purposes
                _logger.LogError(ex, "An error occurred while getting the weather forecasts");

                return StatusCode(500, "An error occurred while processing your request. Please try again later.");
            }
        }

        /// <summary>
        /// Insert a new weather forecast.
        /// </summary>
        /// <param name="forecast">The weather forecast to insert.</param>  
        /// <returns>The inserted weather forecast.</returns>
        /// <response code="201">Returns the newly created weather forecast.</response>
        /// <response code="400">If the weather forecast is null.</response>
        /// <response code="500">If there was an error inserting the weather forecast.</response>
        /// remarks>           
        [HttpPost]
        [ProducesResponseType(typeof(WeatherForecast), 201)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        public async Task<ActionResult> Insert([FromBody] WeatherForecast forecast)
        {
            try
            {
                await _weatherForecastService.InsertForecast(forecast);
                return CreatedAtAction(nameof(Get), new { id = forecast.Id }, forecast);
            }
            catch (Exception)
            {
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Update an existing weather forecast.
        /// </summary>
        /// <param name="id">The id of the weather forecast to update.</param>
        /// <param name="forecast">The updated weather forecast.</param>
        /// <returns></returns>
        /// <response code="204">Returns no content if the weather forecast was updated successfully.</response>
        /// <response code="400">If the weather forecast is null.</response>
        /// <response code="500">If there was an error updating the weather forecast.</response>
        /// remarks>
        [HttpPut("{id}")]
        [ProducesResponseType(204)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        public async Task<ActionResult> Update(string id, [FromBody] WeatherForecast forecast)
        {
            try
            {
                await _weatherForecastService.UpdateForecast(id, forecast);
                return NoContent();
            }
            catch (Exception)
            {
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        ///    Get a weather forecast by id.
        ///    </summary>
        ///    <param name="id">The id of the weather forecast to get.</param>
        ///    <returns>The weather forecast.</returns>
        ///    <response code="200">Returns the weather forecast.</response>
        ///    <response code="404">If the weather forecast is null.</response>
        ///    <response code="500">If there was an error getting the weather forecast.</response>
        ///    remarks>
        [HttpGet("{id}",Name = "GetWeatherForecastById")]
        public async Task<IActionResult> GetForecast(string id)
        {
            var forecast = await _weatherForecastService.GetForecast(id);

            if (forecast == null)
            {
                return NotFound();
            }

            return Ok(forecast);
        }
    }
}