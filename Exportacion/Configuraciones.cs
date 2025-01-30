using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exportacion
{
    internal class Configuraciones
    {
        public string obtenerParametros(string grupo, string subgrupo)
        {
            try
            {
                String ExecutablePath = System.Reflection.Assembly.GetEntryAssembly().Location;

                var builder = new ConfigurationBuilder().SetBasePath(Path.GetDirectoryName(ExecutablePath))
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);


                IConfiguration config = builder.Build();
                return config.GetSection(grupo + ":" + subgrupo).Value.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al obtener el parametro : " + subgrupo + ", " +ex.InnerException.Message + ", " + ex.InnerException.InnerException + "\n");
                throw new Exception("Error al obtener el parametro : " + subgrupo, ex);
                     
            }


            
        }
    }
}
