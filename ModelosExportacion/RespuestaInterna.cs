using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelosExportacion
{
    public class RespuestaInterna
    {
        public RespuestaInterna()
        {
            correcto = true;
            mensaje = "Proceso exitoso!";
            detalle = string.Empty;
            objeto = null;
            horaInicio = DateTime.Now;
        }
        public bool correcto { set; get; }
        public string mensaje { set; get; }
        public string detalle { set; get; }
        public object objeto { set; get; }
        public DataTable tabla { set; get; }
        public DateTime horaInicio { set; get; }
        public DateTime horaFinal { set; get; }

    }
}
