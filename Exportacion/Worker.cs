using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelosExportacion;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exportacion
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            Configuraciones _config = new Configuraciones();
            configuracionesAppSettings config = new configuracionesAppSettings();

            int tiempo = 0;
            //string logPath = string.Empty;


            while (!stoppingToken.IsCancellationRequested)
            {               

                _logger.LogInformation("Inicio de worker: {time}", DateTimeOffset.Now);

                try
                {
                    config.intMiliSegundosEspera = int.Parse(_config.obtenerParametros("Informacion", "MiliSegundosEspera"));
                    tiempo = config.intMiliSegundosEspera;

                    Ejecutar(config, _config);

                }
                catch (Exception ex)
                {
                    _logger.LogError(ex.Message + "\n" + ex.StackTrace);
                }

               
                await Task.Delay(tiempo, stoppingToken);
               

            }
        }


        private async void Ejecutar(configuracionesAppSettings config, Configuraciones _config  )
        {
           

            //  # # # # # # #   VARIABLES DE CONFIGURACIÓN   # # # # # # #  //

            config.bdcnServer = _config.obtenerParametros("ConnectionString", "bdcnServer");
            config.bdcnBD = _config.obtenerParametros("ConnectionString", "bdcnBD");
            config.bdcnUsuario = _config.obtenerParametros("ConnectionString", "bdcnUsuario");
            config.bdcnContraseña = _config.obtenerParametros("ConnectionString", "bdcnContraseña");

            // ICMMNFHeinekenQA_SFTP
            config.sftpcnServer_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpServer");
            config.sftpcnUsuario_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpUsuario");
            config.sftpcnContraseña_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpContraseña");
            config.sftpcnPuerto_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpPuerto");

            //ICMCOMHeinekenQA_SFTP
            config.sftpcnServer_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpServer");
            config.sftpcnUsuario_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpUsuario");
            config.sftpcnContraseña_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpContraseña");
            config.sftpcnPuerto_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpPuerto");

            //ICMCOMCatorcenalHeinekenQA_SFTP
            config.sftpcnServer_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpServer");
            config.sftpcnUsuario_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpUsuario");
            config.sftpcnContraseña_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpContraseña");
            config.sftpcnPuerto_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpPuerto");

            config.strTablasExportar = _config.obtenerParametros("Informacion", "Tablas");
            config.intMaximoRegistros = int.Parse(_config.obtenerParametros("Informacion", "MaximoRegistros"));
            config.strIronLicenseKey = _config.obtenerParametros("Informacion", "IronLicenseKey");
            config.dteFechaFiltro = _config.obtenerParametros("Informacion", "FechaFiltro");

          

            config.ubiRutaCarpetaLocalExcel = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaLocalExcel");
            config.ubiRutaCarpetaDestinoExcel = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaDestinoExcel");




            //  # # # # # # #   INICIA FUNCIONALIDAD   # # # # # # #  //
            ExportarInicio exp = new ExportarInicio(config);
            await exp.ExportarArchivosAsync();

        }

       
    }
}
