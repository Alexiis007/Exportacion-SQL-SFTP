
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelosExportacion
{
    public enum LogType
    {
        INFO,
        ERROR
    }
    internal class Log
    {
        private string _carpetaLogs = string.Empty;
        private string _logSesion = string.Empty;

        //public string Error = "ERROR";
        //public string Info = "INFO";

        public Log()
        {
            this._carpetaLogs = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            verificaCreaCarpetas();
            
            _logSesion = "log_" + DateTime.Today.ToString("yyyyMMdd")+ "_"+ DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + ".log";
        }

        public void verificaCreaCarpetas()
        {
            if (!Directory.Exists(_carpetaLogs))
                Directory.CreateDirectory(_carpetaLogs);
        }

        public void Escribe(LogType tipo, string mensaje, string detalle)
        {
            string texto = formatoTexto(tipo, mensaje, detalle);

            string path = Path.Combine(_carpetaLogs, _logSesion);
            string readText;

            if (File.Exists(path))
            {
                using (StreamReader readtext = new StreamReader(path))
                {
                    readText = readtext.ReadToEnd();
                }
            }
            else
                readText = "";

            using (StreamWriter writetext = new StreamWriter(path))
            {
                writetext.WriteLine(readText + texto);
            }
        }

        private string formatoTexto(LogType tipo, string mensaje, string detalle)
        {
            return tipo + " - " + DateTime.Now + Environment.NewLine +
                    "Mensaje: " + mensaje + Environment.NewLine +
                    "Detalles: " + detalle + Environment.NewLine;
        }
        private string getLogType(int tipo)
        {
            switch (tipo)
            {
                case 0:
                    return "ERROR";
                case 1:
                    return "INFO";
                default:
                    return "OTRO";
            }
        }
    }
}

