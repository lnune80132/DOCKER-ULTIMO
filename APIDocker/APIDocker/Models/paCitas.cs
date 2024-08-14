namespace APIDocker.Models
{
    public class paCitas
    {
        public int IdCita { get; set; }
        public string NombreCliente { get; set; }
        public string TelefonoCliente { get; set; }
        public string EmailCliente { get; set; }
        public string Placa { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public int Ano { get; set; }
        public string NombreSucursal { get; set; }
        public string NombreServicio { get; set; }
        public decimal PrecioServicio { get; set; }
        public DateTime FechaHora { get; set; }
        public string Comentarios { get; set; }

    }
    public class ConfirmacionPaCitas
    {

        public int Codigo { get; set; }
        public string Detalle { get; set; }
        public List<paCitas> Datos { get; set; }
        public object Dato { get; set; }

    }
}
