using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging.Abstractions;
using SFTPService;
using System.Linq.Expressions;
using CsvHelper.Configuration;
using System.Globalization;
using CsvHelper;
using System.Data;
using System.Drawing.Drawing2D;
using SixLabors.ImageSharp.Drawing;




namespace ModelosExportacion
{
    public  class ExportarInicio
    {
        private string bdcnServer;
        private string bdcnBD;
        private string bdcnUsuario;
        private string bdcnContraseña;

        private string sftpServer_ICMCOMHeinekenQA_SFTP;
        private string sftpUsuario_ICMCOMHeinekenQA_SFTP;
        private string sftpContraseña_ICMCOMHeinekenQA_SFTP;
        private int sftpPuerto_ICMCOMHeinekenQA_SFTP;

        private string sftpServer_ICMMNFHeinekenQA_SFTP;
        private string sftpUsuario_ICMMNFHeinekenQA_SFTP;
        private string sftpContraseña_ICMMNFHeinekenQA_SFTP;
        private int sftpPuerto_ICMMNFHeinekenQA_SFTP;

        private string sftpServer_ICMCOMCatorcenalHeinekenQA_SFTP;
        private string sftpUsuario_ICMCOMCatorcenalHeinekenQA_SFTP;
        private string sftpContraseña_ICMCOMCatorcenalHeinekenQA_SFTP;
        private int sftpPuerto_ICMCOMCatorcenalHeinekenQA_SFTP;

        private string strTablasExportar;
        private int intMaximoRegistros;      
        private string dteFechaFiltro;

        private string ubiRutaCarpetaLocal;
        private string ubiRutaCarpetaDestino;

        private string ubiRutaCarpetaDestinoTablasToSQL;
        private string ubiRutaCarpetaOrigenTablasToSQL;

        private SftpConfig ICMMNFHeinekenQA_SFTP;
        private SftpConfig ICMCOMHeinekenQA_SFTP;
        private SftpConfig ICMCOMCatorcenalHeinekenQA_SFTP;

        private Log log;
        Encoder encoder = new Encoder();        

        public ExportarInicio(configuracionesAppSettings config)
        {
            //Informacion SQL-BD donde se extrae la data para varicent
            this.bdcnServer = encoder.DesEncriptarBase64(config.bdcnServer);
            this.bdcnBD = encoder.DesEncriptarBase64(config.bdcnBD);
            this.bdcnUsuario = encoder.DesEncriptarBase64(config.bdcnUsuario);
            this.bdcnContraseña = encoder.DesEncriptarBase64(config.bdcnContraseña);

            //Ubicacion local de las tablas y ubicacion destino de los modelos de varicent (SQL to Varicent)
            this.ubiRutaCarpetaLocal = config.ubiRutaCarpetaLocalExcel;
            this.ubiRutaCarpetaDestino = config.ubiRutaCarpetaDestinoExcel;

            //Tablas que se necesitan extraer de SQL-BD para varicent
            this.strTablasExportar = config.strTablasExportar;    

            // ICMMNFHeinekenQA_SFTP
            this.sftpServer_ICMMNFHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnServer_ICMMNFHeinekenQA_SFTP); 
            this.sftpUsuario_ICMMNFHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnUsuario_ICMMNFHeinekenQA_SFTP); 
            this.sftpContraseña_ICMMNFHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnContraseña_ICMMNFHeinekenQA_SFTP); 
            this.sftpPuerto_ICMMNFHeinekenQA_SFTP =   int.Parse(encoder.DesEncriptarBase64(config.sftpcnPuerto_ICMMNFHeinekenQA_SFTP));

            // ICMCOMHeinekenQA_SFTP
            this.sftpServer_ICMCOMHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnServer_ICMCOMHeinekenQA_SFTP);
            this.sftpUsuario_ICMCOMHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnUsuario_ICMCOMHeinekenQA_SFTP);
            this.sftpContraseña_ICMCOMHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnContraseña_ICMCOMHeinekenQA_SFTP);
            this.sftpPuerto_ICMCOMHeinekenQA_SFTP = int.Parse(encoder.DesEncriptarBase64(config.sftpcnPuerto_ICMCOMHeinekenQA_SFTP));

            // ICMCOMCatorcenalHeinekenQA_SFTP
            this.sftpServer_ICMCOMCatorcenalHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnServer_ICMCOMCatorcenalHeinekenQA_SFTP);
            this.sftpUsuario_ICMCOMCatorcenalHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnUsuario_ICMCOMCatorcenalHeinekenQA_SFTP);
            this.sftpContraseña_ICMCOMCatorcenalHeinekenQA_SFTP = encoder.DesEncriptarBase64(config.sftpcnContraseña_ICMCOMCatorcenalHeinekenQA_SFTP);
            this.sftpPuerto_ICMCOMCatorcenalHeinekenQA_SFTP = int.Parse(encoder.DesEncriptarBase64(config.sftpcnPuerto_ICMCOMCatorcenalHeinekenQA_SFTP));

            this.intMaximoRegistros = config.intMaximoRegistros;
            this.dteFechaFiltro = config.dteFechaFiltro;

            //Ubicacion remota de las tablas y ubicacion destino de la carpeta contenedora local (Varicent to SQL)
            this.ubiRutaCarpetaDestinoTablasToSQL = config.ubiRutaCarpetaDestinoTablasToSQL;
            this.ubiRutaCarpetaOrigenTablasToSQL = config.ubiRutaCarpetaOrigenTablasToSQL;

            ICMMNFHeinekenQA_SFTP = new SftpConfig
            {
                Host = sftpServer_ICMMNFHeinekenQA_SFTP,
                Port = sftpPuerto_ICMMNFHeinekenQA_SFTP,
                UserName = sftpUsuario_ICMMNFHeinekenQA_SFTP,
                Password = sftpContraseña_ICMMNFHeinekenQA_SFTP
            };

            ICMCOMHeinekenQA_SFTP = new SftpConfig
            {
                Host = sftpServer_ICMCOMHeinekenQA_SFTP,
                Port = sftpPuerto_ICMCOMHeinekenQA_SFTP,
                UserName = sftpUsuario_ICMCOMHeinekenQA_SFTP,
                Password = sftpContraseña_ICMCOMHeinekenQA_SFTP
            };

            ICMCOMCatorcenalHeinekenQA_SFTP = new SftpConfig
            {
                Host = sftpServer_ICMCOMCatorcenalHeinekenQA_SFTP,
                Port = sftpPuerto_ICMCOMCatorcenalHeinekenQA_SFTP,
                UserName = sftpUsuario_ICMCOMCatorcenalHeinekenQA_SFTP,
                Password = sftpContraseña_ICMCOMCatorcenalHeinekenQA_SFTP
            };

            log = new Log();
        }
  
        private async Task<RespuestaInterna> ExportarArchivo(string tabla, ConexionBD bd, string RutaExcel, string query)
        {
          
            RespuestaInterna exportacion = new RespuestaInterna();
            Task<RespuestaInterna> respuesta = bd.ejecutScript(query);

            if (respuesta.Result.correcto)
            {

                //string contieneBegda = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TableName AND COLUMN_NAME = @ColumnName";

                Console.WriteLine("Exportando tabla");
                log.Escribe(LogType.INFO, "Exportando tabla", "Exportacion de Informacion");

                exportacion = await Funciones.ExportCSV(respuesta.Result.tabla, tabla, RutaExcel);

                log.Escribe(LogType.INFO, exportacion.mensaje, exportacion.detalle);
            }
            else
            {
                exportacion.correcto = false;
                exportacion.mensaje = respuesta.Result.mensaje;
                exportacion.detalle = respuesta.Result.detalle;

                Console.WriteLine("Error al consultar la tabla: " + respuesta.Result.mensaje);
                log.Escribe(LogType.ERROR, respuesta.Result.mensaje, "");
            }

            return exportacion;
        }
       
        private RespuestaInterna LimpiarDirectorio(string CarpetaOrigen, string tipoArchivos=".csv")
        {
            RespuestaInterna limpiado = new RespuestaInterna();

            string[] files = Directory.GetFiles(CarpetaOrigen, "*"+tipoArchivos, SearchOption.AllDirectories);
            string mensaje = "";
            try
            {
                foreach (string item in files)
                {

                    File.Delete(item);
                }

                limpiado.correcto = true;
                limpiado.mensaje = "Limpieza del directorio " + CarpetaOrigen + " exitosa";
            }
            catch (Exception ex)
            {
                mensaje = "Error al hacer la limpieza del directorio " + CarpetaOrigen + "\n" + ex.Message;
               
                limpiado.correcto = false;
                limpiado.mensaje= mensaje;
                limpiado.detalle = ex.StackTrace;
                
            }
          
            return limpiado;
        }

        public async Task ExportarArchivosAsyncToSFTP()
        {
            Console.Clear();

            Console.WriteLine("       Inicia Proceso de Exportacion!");
            Console.WriteLine("Fecha Inicio:  " + DateTime.Now.ToString());
            Console.WriteLine("==============================================================");

            log.Escribe(LogType.INFO, "Inicia Proceso de Exportacion!", "Exportar archivos");

            ConexionBD bd = new ConexionBD(bdcnServer, bdcnBD, bdcnUsuario, bdcnContraseña);
            bool prueba = await bd.probarConexion();

            if (prueba)
            {
                Console.WriteLine(" == Prueba de conexion BD Exitosa == ");
                log.Escribe(LogType.INFO, "== Prueba de conexion BD Exitosa == ", "Se conecto a SQL de manera correcta");

                RespuestaInterna exportacion_limpieza = new RespuestaInterna();
                exportacion_limpieza = LimpiarDirectorio(ubiRutaCarpetaLocal);
                
                if (exportacion_limpieza.correcto)
                {
                    log.Escribe(LogType.INFO, exportacion_limpieza.mensaje, exportacion_limpieza.detalle);

                    String[] tablas = strTablasExportar.Split(',');

                    foreach (String tabla in tablas)
                    {
                        string NombreTabla = tabla.Trim();
                        string mensaje = "";
                        string where = " Where BEGDA >= '" + this.dteFechaFiltro + "'";
                        string RutaCSV = ubiRutaCarpetaLocal + "\\" + NombreTabla + ".csv";
                        string query = "Select * from " + tabla.Trim();

                        Console.WriteLine("===========================================");
                        Console.WriteLine("      " + tabla);
                        Console.WriteLine("===========================================");
                        log.Escribe(LogType.INFO, " =========================================== \n" + tabla + "\n =========================================== ", "Procesando la tabla");

                        // Se manda a ejecutar el query para la obtener la informacion de la tabla                        
                        RespuestaInterna exportacion = new RespuestaInterna();
                        exportacion = ExportarArchivo(tabla.Trim(), bd, RutaCSV, query).Result;

                        Console.WriteLine("===========================================");
                        log.Escribe(LogType.INFO, " ===========================================", "");
                    }

                    RespuestaInterna ICMMNFHeinekenQA = new RespuestaInterna();
                    RespuestaInterna ICMCOMHeinekenQA = new RespuestaInterna();
                    RespuestaInterna ICMCOMCatorcenalHeinekenQA = new RespuestaInterna();

                    ICMMNFHeinekenQA = await Funciones.ExportarFTP(ubiRutaCarpetaLocal, ubiRutaCarpetaDestino, ICMMNFHeinekenQA_SFTP, "ICMMNFHeinekenQA_Tablas");
                    ICMCOMHeinekenQA = await Funciones.ExportarFTP(ubiRutaCarpetaLocal, ubiRutaCarpetaDestino, ICMCOMHeinekenQA_SFTP, "ICMCOMHeinekenQA_Tablas");
                    ICMCOMCatorcenalHeinekenQA = await Funciones.ExportarFTP(ubiRutaCarpetaLocal, ubiRutaCarpetaDestino, ICMCOMCatorcenalHeinekenQA_SFTP, "ICMCOMCatorcenalHeinekenQA_Tablas");


                    Console.WriteLine(ICMMNFHeinekenQA.mensaje);
                    log.Escribe(ICMMNFHeinekenQA.correcto ? LogType.INFO : LogType.ERROR, ICMMNFHeinekenQA.mensaje, ICMMNFHeinekenQA.detalle);
                    Console.WriteLine("===========================================");
                    Console.WriteLine(ICMCOMHeinekenQA.mensaje);
                    log.Escribe(ICMCOMHeinekenQA.correcto ? LogType.INFO : LogType.ERROR, ICMCOMHeinekenQA.mensaje, ICMCOMHeinekenQA.detalle);
                    Console.WriteLine("===========================================");
                    Console.WriteLine(ICMCOMCatorcenalHeinekenQA.mensaje);
                    log.Escribe(ICMCOMCatorcenalHeinekenQA.correcto ? LogType.INFO : LogType.ERROR, ICMCOMCatorcenalHeinekenQA.mensaje, ICMCOMCatorcenalHeinekenQA.detalle);
                    Console.WriteLine("===========================================");
                }
                else
                {
                    Console.WriteLine("Error al hacer la limpieza al directorio Origen");
                    log.Escribe(LogType.ERROR, "Error limpieza directorio", exportacion_limpieza.mensaje);
                }
            }
            else
            {
                Console.WriteLine("Error al conectarse a la BD");
                log.Escribe(LogType.ERROR, "Error al conectarse a la BD", "");
            }

           
           Console.WriteLine("Fecha Termino:  " + DateTime.Now.ToString());
           log.Escribe(LogType.INFO, "Fecha Termino:  " + DateTime.Now.ToString(), "Proceso Terminado");

        }

        public void ExportarArchivosAsyncToSQL()
        {
            Console.Clear();

            Console.WriteLine("       Inicia Proceso de Exportacion a SQL!");
            Console.WriteLine("Fecha Inicio:  " + DateTime.Now.ToString());
            Console.WriteLine("==============================================================");

            log.Escribe(LogType.INFO, "Inicia Proceso de Exportacion!", "Exportar Data");

            LimpiarDirectorio(ubiRutaCarpetaDestinoTablasToSQL, ".csv");

            // Descarga archivos de SFTP Modelo Manufactura
            Funciones.DownloadCSV(ubiRutaCarpetaOrigenTablasToSQL, ubiRutaCarpetaDestinoTablasToSQL, ICMMNFHeinekenQA_SFTP, "TablasManufacturaToSQL");
            // Descarga archivos de SFTP Modelo Quincenal
            Funciones.DownloadCSV(ubiRutaCarpetaOrigenTablasToSQL, ubiRutaCarpetaDestinoTablasToSQL, ICMCOMHeinekenQA_SFTP, "TablasQuincenalToSQL");
            // Descarga archivos de SFTP Modelo Catorcenal
            Funciones.DownloadCSV(ubiRutaCarpetaOrigenTablasToSQL, ubiRutaCarpetaDestinoTablasToSQL, ICMCOMCatorcenalHeinekenQA_SFTP, "TablasCatorcenalToSQL");

            Console.WriteLine("==============================================================");
            Console.WriteLine("Procesando los archivos descargados para la exportacion a SQL.");
            Console.WriteLine("==============================================================");

            // Se procesan CSV para despues ser migrados por BulkCopy a la BD de SQL
            Funciones.CSVToQuery(ubiRutaCarpetaDestinoTablasToSQL, bdcnServer, bdcnBD, bdcnUsuario, bdcnContraseña);
        }
    }
}
