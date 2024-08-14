using Microsoft.AspNetCore.Mvc;
using ClosedXML.Excel;
using System.IO;
using System.Text.Json;
using APIDocker.Models;

[Route("api/[controller]")]
[ApiController]
public class ReporteController : ControllerBase
{
    private readonly IHttpClientFactory _clientFactory;

    public ReporteController(IHttpClientFactory clientFactory)
    {
        _clientFactory = clientFactory;
    }

    // ===============================================================================
    // TODAS

    [HttpGet]
    [Route("conexionDocker/TotalCitas")]
    public async Task<IActionResult> ConexionDocker()
    {
        try
        {
            // Consumir el API para obtener las citas
            var client = _clientFactory.CreateClient();
            var response = await client.GetAsync("https://grupomotoresbritanicos.somee.com/Citas/TodasCitas");

            if (!response.IsSuccessStatusCode)
            {
                return StatusCode((int)response.StatusCode, "Error al obtener datos del API");
            }

            var responseString = await response.Content.ReadAsStringAsync();
            var confirmacionPaCitas = JsonSerializer.Deserialize<ConfirmacionPaCitas>(responseString, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (confirmacionPaCitas == null || confirmacionPaCitas.Codigo != 0)
            {
                return StatusCode(500, "Error en la respuesta del API");
            }

            var citas = confirmacionPaCitas.Datos;

            // Usar MemoryStream para almacenar el archivo Excel
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("CitasTotales");


                // Añadir cabeceras
                worksheet.Cell(2, 1).Value = "IdCita";
                worksheet.Cell(2, 2).Value = "NombreCliente";
                worksheet.Cell(2, 3).Value = "TelefonoCliente";
                worksheet.Cell(2, 4).Value = "EmailCliente";
                worksheet.Cell(2, 5).Value = "Placa";
                worksheet.Cell(2, 6).Value = "Marca";
                worksheet.Cell(2, 7).Value = "Modelo";
                worksheet.Cell(2, 8).Value = "Ano";
                worksheet.Cell(2, 9).Value = "NombreSucursal";
                worksheet.Cell(2, 10).Value = "NombreServicio";
                worksheet.Cell(2, 11).Value = "PrecioServicio";
                worksheet.Cell(2, 12).Value = "FechaHora";
                worksheet.Cell(2, 13).Value = "Comentarios";
                // Ajustar el ancho de las columnas automáticamente
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }

                // Añadir datos dinámicamente
                int currentRow = 3;
                foreach (var cita in citas)
                {
                    worksheet.Cell(currentRow, 1).Value = cita.IdCita;
                    worksheet.Cell(currentRow, 2).Value = cita.NombreCliente;
                    worksheet.Cell(currentRow, 3).Value = cita.TelefonoCliente;
                    worksheet.Cell(currentRow, 4).Value = cita.EmailCliente;
                    worksheet.Cell(currentRow, 5).Value = cita.Placa;
                    worksheet.Cell(currentRow, 6).Value = cita.Marca;
                    worksheet.Cell(currentRow, 7).Value = cita.Modelo;
                    worksheet.Cell(currentRow, 8).Value = cita.Ano;
                    worksheet.Cell(currentRow, 9).Value = cita.NombreSucursal;
                    worksheet.Cell(currentRow, 10).Value = cita.NombreServicio;
                    worksheet.Cell(currentRow, 11).Value = cita.PrecioServicio;
                    worksheet.Cell(currentRow, 12).Value = cita.FechaHora;
                    worksheet.Cell(currentRow, 13).Value = cita.Comentarios;
                    currentRow++;
                }
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }
                // Crear un MemoryStream y guardar el archivo Excel en él
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0; // Reiniciar la posición del stream a 0

                    // Devolver el archivo como una descarga
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CitasTotales.xlsx");
                }
            }
        }
        catch (Exception ex)
        {
            // Manejo de errores, por ejemplo, logueo de excepciones
            // _logger.LogError(ex, "Ocurrió un error al generar el archivo Excel.");

            // Devuelve un error genérico
            return StatusCode(500, "Ocurrió un error interno al procesar tu solicitud.");
        }
    }

    // ===============================================================================
    // SEGUN SUCURSALES

    [HttpGet]
    [Route("conexionDocker/TotalCitasPorSucursal")]
    public async Task<IActionResult> ConexionDocker(long idSucursal)
    {
        try
        {
            // Consumir el API para obtener las citas por sucursal
            var client = _clientFactory.CreateClient();
            var response = await client.GetAsync($"https://grupomotoresbritanicos.somee.com/Citas/ConsultarCitaPorSucursal?idSucursal={idSucursal}");

            if (!response.IsSuccessStatusCode)
            {
                return StatusCode((int)response.StatusCode, "Error al obtener datos del API");
            }

            var responseString = await response.Content.ReadAsStringAsync();
            var confirmacionPaCitas = JsonSerializer.Deserialize<ConfirmacionPaCitas>(responseString, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (confirmacionPaCitas == null || confirmacionPaCitas.Codigo != 0)
            {
                return StatusCode(500, "Error en la respuesta del API");
            }

            var citas = confirmacionPaCitas.Datos;

            // Usar MemoryStream para almacenar el archivo Excel
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("CitasPorSucursal");

                // Añadir cabeceras
                worksheet.Cell(2, 1).Value = "IdCita";
                worksheet.Cell(2, 2).Value = "NombreCliente";
                worksheet.Cell(2, 3).Value = "TelefonoCliente";
                worksheet.Cell(2, 4).Value = "EmailCliente";
                worksheet.Cell(2, 5).Value = "Placa";
                worksheet.Cell(2, 6).Value = "Marca";
                worksheet.Cell(2, 7).Value = "Modelo";
                worksheet.Cell(2, 8).Value = "Ano";
                worksheet.Cell(2, 9).Value = "NombreSucursal";
                worksheet.Cell(2, 10).Value = "NombreServicio";
                worksheet.Cell(2, 11).Value = "PrecioServicio";
                worksheet.Cell(2, 12).Value = "FechaHora";
                worksheet.Cell(2, 13).Value = "Comentarios";
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }
                // Añadir datos dinámicamente
                int currentRow = 3;
                foreach (var cita in citas)
                {
                    worksheet.Cell(currentRow, 1).Value = cita.IdCita;
                    worksheet.Cell(currentRow, 2).Value = cita.NombreCliente;
                    worksheet.Cell(currentRow, 3).Value = cita.TelefonoCliente;
                    worksheet.Cell(currentRow, 4).Value = cita.EmailCliente;
                    worksheet.Cell(currentRow, 5).Value = cita.Placa;
                    worksheet.Cell(currentRow, 6).Value = cita.Marca;
                    worksheet.Cell(currentRow, 7).Value = cita.Modelo;
                    worksheet.Cell(currentRow, 8).Value = cita.Ano;
                    worksheet.Cell(currentRow, 9).Value = cita.NombreSucursal;
                    worksheet.Cell(currentRow, 10).Value = cita.NombreServicio;
                    worksheet.Cell(currentRow, 11).Value = cita.PrecioServicio;
                    worksheet.Cell(currentRow, 12).Value = cita.FechaHora;
                    worksheet.Cell(currentRow, 13).Value = cita.Comentarios;
                    currentRow++;
                }
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }
                // Crear un MemoryStream y guardar el archivo Excel en él
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0; // Reiniciar la posición del stream a 0

                    // Devolver el archivo como una descarga
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CitasPorSucursal.xlsx");
                }
            }
        }
        catch (Exception ex)
        {
            // Manejo de errores, por ejemplo, logueo de excepciones
            // _logger.LogError(ex, "Ocurrió un error al generar el archivo Excel.");

            // Devuelve un error genérico
            return StatusCode(500, "Ocurrió un error interno al procesar tu solicitud.");
        }
    }

    // ===============================================================================
    // SEGUN MES

    [HttpGet]
    [Route("conexionDocker/TotalCitasPorMes")]
    public async Task<IActionResult> TotalCitasPorMes(int mes)
    {
        try
        {
            var client = _clientFactory.CreateClient();

            if (mes < 1 || mes > 12)
            {
                return BadRequest("El mes proporcionado no es válido.");
            }

            var fecha = new DateTime(2024, mes, 1).ToString("yyyy-MM-dd");

            var response = await client.GetAsync($"https://grupomotoresbritanicos.somee.com/Citas/ConsultarCitaPorMes?fecha={fecha:yyyy-MM-dd}");

            if (!response.IsSuccessStatusCode)
            {
                return StatusCode((int)response.StatusCode, "Error al obtener datos del API");
            }

            var responseString = await response.Content.ReadAsStringAsync();
            var confirmacionPaCitas = JsonSerializer.Deserialize<ConfirmacionPaCitas>(responseString, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (confirmacionPaCitas == null || confirmacionPaCitas.Codigo != 0)
            {
                return StatusCode(500, "Error en la respuesta del API");
            }

            var citas = confirmacionPaCitas.Datos;

            // Usar MemoryStream para almacenar el archivo Excel
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("CitasPorMes");

                // Añadir cabeceras
                worksheet.Cell(2, 1).Value = "IdCita";
                worksheet.Cell(2, 2).Value = "NombreCliente";
                worksheet.Cell(2, 3).Value = "TelefonoCliente";
                worksheet.Cell(2, 4).Value = "EmailCliente";
                worksheet.Cell(2, 5).Value = "Placa";
                worksheet.Cell(2, 6).Value = "Marca";
                worksheet.Cell(2, 7).Value = "Modelo";
                worksheet.Cell(2, 8).Value = "Ano";
                worksheet.Cell(2, 9).Value = "NombreSucursal";
                worksheet.Cell(2, 10).Value = "NombreServicio";
                worksheet.Cell(2, 11).Value = "PrecioServicio";
                worksheet.Cell(2, 12).Value = "FechaHora";
                worksheet.Cell(2, 13).Value = "Comentarios";
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }
                // Añadir datos dinámicamente
                int currentRow = 3;
                foreach (var cita in citas)
                {
                    worksheet.Cell(currentRow, 1).Value = cita.IdCita;
                    worksheet.Cell(currentRow, 2).Value = cita.NombreCliente;
                    worksheet.Cell(currentRow, 3).Value = cita.TelefonoCliente;
                    worksheet.Cell(currentRow, 4).Value = cita.EmailCliente;
                    worksheet.Cell(currentRow, 5).Value = cita.Placa;
                    worksheet.Cell(currentRow, 6).Value = cita.Marca;
                    worksheet.Cell(currentRow, 7).Value = cita.Modelo;
                    worksheet.Cell(currentRow, 8).Value = cita.Ano;
                    worksheet.Cell(currentRow, 9).Value = cita.NombreSucursal;
                    worksheet.Cell(currentRow, 10).Value = cita.NombreServicio;
                    worksheet.Cell(currentRow, 11).Value = cita.PrecioServicio;
                    worksheet.Cell(currentRow, 12).Value = cita.FechaHora;
                    worksheet.Cell(currentRow, 13).Value = cita.Comentarios;
                    currentRow++;
                }
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }

                // Crear un MemoryStream y guardar el archivo Excel en él
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0; // Reiniciar la posición del stream a 0

                    // Devolver el archivo como una descarga
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CitasPorMes.xlsx");
                }
            }
        }
        catch (Exception ex)
        {
            // Manejo de errores, por ejemplo, logueo de excepciones
            // _logger.LogError(ex, "Ocurrió un error al generar el archivo Excel.");

            // Devuelve un error genérico
            return StatusCode(500, "Ocurrió un error interno al procesar tu solicitud.");
        }
    }

    // ===============================================================================
    // SEGUN FECHA

    [HttpGet]
    [Route("conexionDocker/TotalCitasPorSemana")]
    public async Task<IActionResult> TotalCitasPorSemana(DateTime fecha)
    {
        try
        {
            // Consumir el API para obtener las citas por semana
            var client = _clientFactory.CreateClient();

            var response = await client.GetAsync($"https://grupomotoresbritanicos.somee.com/Citas/ConsultarCitaPorSemana?fecha={fecha:yyyy-MM-dd}");

            if (!response.IsSuccessStatusCode)
            {
                return StatusCode((int)response.StatusCode, "Error al obtener datos del API");
            }

            var responseString = await response.Content.ReadAsStringAsync();
            var confirmacionPaCitas = JsonSerializer.Deserialize<ConfirmacionPaCitas>(responseString, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (confirmacionPaCitas == null || confirmacionPaCitas.Codigo != 0)
            {
                return StatusCode(500, "Error en la respuesta del API");
            }

            var citas = confirmacionPaCitas.Datos;

            // Usar MemoryStream para almacenar el archivo Excel
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("CitasPorSemana");

                // Añadir cabeceras
                worksheet.Cell(2, 1).Value = "IdCita";
                worksheet.Cell(2, 2).Value = "NombreCliente";
                worksheet.Cell(2, 3).Value = "TelefonoCliente";
                worksheet.Cell(2, 4).Value = "EmailCliente";
                worksheet.Cell(2, 5).Value = "Placa";
                worksheet.Cell(2, 6).Value = "Marca";
                worksheet.Cell(2, 7).Value = "Modelo";
                worksheet.Cell(2, 8).Value = "Ano";
                worksheet.Cell(2, 9).Value = "NombreSucursal";
                worksheet.Cell(2, 10).Value = "NombreServicio";
                worksheet.Cell(2, 11).Value = "PrecioServicio";
                worksheet.Cell(2, 12).Value = "FechaHora";
                worksheet.Cell(2, 13).Value = "Comentarios";
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }
                // Añadir datos dinámicamente
                int currentRow = 3;
                foreach (var cita in citas)
                {
                    worksheet.Cell(currentRow, 1).Value = cita.IdCita;
                    worksheet.Cell(currentRow, 2).Value = cita.NombreCliente;
                    worksheet.Cell(currentRow, 3).Value = cita.TelefonoCliente;
                    worksheet.Cell(currentRow, 4).Value = cita.EmailCliente;
                    worksheet.Cell(currentRow, 5).Value = cita.Placa;
                    worksheet.Cell(currentRow, 6).Value = cita.Marca;
                    worksheet.Cell(currentRow, 7).Value = cita.Modelo;
                    worksheet.Cell(currentRow, 8).Value = cita.Ano;
                    worksheet.Cell(currentRow, 9).Value = cita.NombreSucursal;
                    worksheet.Cell(currentRow, 10).Value = cita.NombreServicio;
                    worksheet.Cell(currentRow, 11).Value = cita.PrecioServicio;
                    worksheet.Cell(currentRow, 12).Value = cita.FechaHora;
                    worksheet.Cell(currentRow, 13).Value = cita.Comentarios;
                    currentRow++;
                }
                // Ajustar el ancho de las columnas después de añadir los datos
                for (int i = 1; i <= 13; i++)
                {
                    worksheet.Column(i).AdjustToContents();
                }

                // Crear un MemoryStream y guardar el archivo Excel en él
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0; // Reiniciar la posición del stream a 0

                    // Devolver el archivo como una descarga
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CitasPorSemana.xlsx");
                }
            }
        }
        catch (Exception ex)
        {
            // Manejo de errores, por ejemplo, logueo de excepciones
            // _logger.LogError(ex, "Ocurrió un error al generar el archivo Excel.");

            // Devuelve un error genérico
            return StatusCode(500, "Ocurrió un error interno al procesar tu solicitud.");
        }
    }
}