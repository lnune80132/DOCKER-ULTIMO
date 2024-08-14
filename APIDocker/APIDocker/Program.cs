using System.Security.Cryptography.X509Certificates;

var builder = WebApplication.CreateBuilder(args);

// Obtener el puerto del entorno o usar el 5024 como predeterminado
var port = Environment.GetEnvironmentVariable("PORT") ?? "5024";
builder.WebHost.UseUrls($"http://*:{port}");

// Add services to the container.
builder.Services.AddControllers();

// Add IHttpClientFactory
builder.Services.AddHttpClient();

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment() || app.Environment.IsProduction())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "My API V1");
        c.RoutePrefix = string.Empty; // Configura Swagger para estar en la raíz de la aplicación
    });
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
