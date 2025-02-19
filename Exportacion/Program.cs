// See https://aka.ms/new-console-template for more information
using Microsoft.Extensions.Hosting;
using ModelosExportacion;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using CsvHelper;
using System.Runtime.Serialization;


namespace Exportacion
{
    class Program
    {
        static async Task Main(string[] args)
        {                 
            Configuraciones _config = new Configuraciones();
            configuracionesAppSettings config = new configuracionesAppSettings();
           
            //  # # # # # # #   VARIABLES DE CONFIGURACIÓN   # # # # # # #  //

            //Informacion SQL-BD donde se extrae la data para varicent
            config.bdcnServer = _config.obtenerParametros("ConnectionString", "bdcnServer");
            config.bdcnBD = _config.obtenerParametros("ConnectionString", "bdcnBD");
            config.bdcnUsuario = _config.obtenerParametros("ConnectionString", "bdcnUsuario");
            config.bdcnContraseña = _config.obtenerParametros("ConnectionString", "bdcnContraseña");

            // Modelo 1 ICMMNFHeinekenQA_SFTP
            config.sftpcnServer_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpServer");
            config.sftpcnUsuario_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpUsuario");
            config.sftpcnContraseña_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpContraseña");
            config.sftpcnPuerto_ICMMNFHeinekenQA_SFTP = _config.obtenerParametros("ICMMNFHeinekenQA_SFTP", "sftpPuerto");

            // Modelo 2 ICMCOMHeinekenQA_SFTP
            config.sftpcnServer_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpServer");
            config.sftpcnUsuario_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpUsuario");
            config.sftpcnContraseña_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpContraseña");
            config.sftpcnPuerto_ICMCOMHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMHeinekenQA_SFTP", "sftpPuerto");

            // Modelo 3 ICMCOMCatorcenalHeinekenQA_SFTP
            config.sftpcnServer_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpServer");
            config.sftpcnUsuario_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpUsuario");
            config.sftpcnContraseña_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpContraseña");
            config.sftpcnPuerto_ICMCOMCatorcenalHeinekenQA_SFTP = _config.obtenerParametros("ICMCOMCatorcenalHeinekenQA_SFTP", "sftpPuerto");

            //Tablas que se necesitan extraer de SQL-BD para varicent
            config.strTablasExportar = _config.obtenerParametros("Informacion", "Tablas");
            config.intMaximoRegistros = int.Parse(_config.obtenerParametros("Informacion", "MaximoRegistros"));
            config.dteFechaFiltro = _config.obtenerParametros("Informacion", "FechaFiltro");

            //Ubicacion local de las tablas y ubicacion destino de los modelos de varicent (SQL to Varicent)
            config.ubiRutaCarpetaLocalExcel = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaLocalExcel");
            config.ubiRutaCarpetaDestinoExcel = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaDestinoExcel");

            //Ubicacion remota de las tablas y ubicacion destino de la carpeta contenedora local (Varicent to SQL)
            config.ubiRutaCarpetaDestinoTablasToSQL = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaLocalResultModels");
            config.ubiRutaCarpetaOrigenTablasToSQL = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaOrigenResultModels");
            
            ExportarInicio exp = new ExportarInicio(config);

            // Proceso Exportacion SQL To Varicent
            await exp.ExportarArchivosAsyncToSFTP();

            // Proceso Exportacion Varicent To SQL
<<<<<<< HEAD
            exp.ExportarArchivosAsyncToSQL();
=======
            //await exp.ExportarArchivosAsyncToSQL();
>>>>>>> 110a93f0ddff59581a9aad38d4846cb77ac68682
        }
    
    }
}

