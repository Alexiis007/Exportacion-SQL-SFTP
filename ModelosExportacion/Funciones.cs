using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using IronXL;
using System.Diagnostics;
using SixLabors.Fonts.Tables.AdvancedTypographic;
using SFTPService;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Configuration;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using Microsoft.SqlServer.Server;
using System.Drawing;
using Renci.SshNet;
using System.Reflection;
using System.Collections;
using System.Data.Common;
using static OfficeOpenXml.ExcelErrorValue;
using Microsoft.VisualBasic.FileIO;
using static Org.BouncyCastle.Crypto.Engines.SM2Engine;



namespace ModelosExportacion
{
    static class Funciones
    {
        // Exporta tablas SQL a CSVs a la ruta local
        public static async Task<RespuestaInterna> ExportCSV(System.Data.DataTable tabla, string NombreTabla, string RutaDestino, string rutaLocal)
        {
            string mensaje = "";
            RespuestaInterna respInt = new RespuestaInterna();

            if (tabla != null && tabla.Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(RutaDestino))
                {
                    try
                    {
                        // ------ Creacion de CSVs ------
                        using (StreamWriter sw = new StreamWriter(RutaDestino))
                        {
                            StringBuilder sb = new StringBuilder();

                            string[] NombresColumnas = tabla.Columns.Cast<DataColumn>().Select(columna => columna.ColumnName).ToArray();
                            string encabezados = string.Join(",", NombresColumnas);
                            sb.AppendLine(encabezados);

                            List<string> valoresxlinea = tabla.AsEnumerable().Select(row => string.Join(",", row.ItemArray.Select(valor =>
                            {
                                string valorString = valor.ToString();

                                // Si el valor contiene coma, salto de línea o comillas dobles, rodearlo con comillas dobles
                                if (valorString.Contains(",") || valorString.Contains("\n") || valorString.Contains("\""))
                                {                                    
                                    valorString = "\"" + valorString.Replace("\"", "\"\"") + "\"";
                                }

                                return valorString;
                            }))).ToList();
                            string rows = string.Join(Environment.NewLine, valoresxlinea);
                            sb.AppendLine(rows);

                            await sw.WriteLineAsync(sb.ToString());
                            sw.Close();
                            sb.Clear();
                        }

                        // ----- Configuraciones De CSVs ------
                        string[] ArchivosCSV = Directory.GetFiles(rutaLocal, "*.csv");
                        var reader = new StreamReader(RutaDestino);
                        var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));
                        var filas = csv.GetRecords<dynamic>().ToList();

                        foreach (var fila in filas)
                        {
                            var diccionarioFila = fila as IDictionary<string, object>;
                            if (diccionarioFila != null)
                            {
                                foreach (var col in diccionarioFila)
                                {
                                    // Formateamos la fecha a dd/MM/yyyy
                                    if (col.Value is string colValue && DateTime.TryParseExact(colValue, "dd/MM/yyyy hh:mm:ss tt", new CultureInfo("es-ES"), DateTimeStyles.None, out DateTime fecha))
                                    {
                                        diccionarioFila[col.Key] = fecha.ToString("dd/MM/yyyy");

                                        // Si el año pasa del 9998 se cambia a 9998 
                                        int año = int.Parse(diccionarioFila[col.Key].ToString().Split("/")[2]);
                                        if (año > 9998)
                                        {
                                            DateTime fechaPredeterminada = new DateTime(9998, 1, 1); // Fecha predeterminada
                                            diccionarioFila[col.Key] = fechaPredeterminada.ToString("dd/MM/yyyy");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("El archivo "+ NombreTabla +" no contiene informacion");
                            }
                        }
                        reader.Close();

                        //----- Guardamos el archivo con sus modificaciones hechas CSV -----                        
                        var writer = new StreamWriter(RutaDestino);                       
                        var CsvWriter = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture));                        
                        CsvWriter.WriteRecords(filas);
                        writer.Close();

                        mensaje = "ExportToCSV: Tabla '" + NombreTabla + "' exportada a csv con " + tabla.Rows.Count.ToString() + " filas";
                        Console.WriteLine(mensaje);
                        respInt.correcto = true;
                        respInt.mensaje = mensaje;
                        respInt.detalle = "";
                    }
                    catch (Exception ex)
                    {

                        mensaje = "ExportToCSV: Error - La tabla '" + NombreTabla + "'.\n" + ex.Message;
                        Console.WriteLine(mensaje);

                        respInt.correcto = false;
                        respInt.mensaje = mensaje;
                        respInt.detalle = ex.StackTrace;
                    }

                }
                else
                {
                    mensaje = "ExportToCSV: La ruta " + RutaDestino + " No es valida!";
                    Console.WriteLine(mensaje);

                    respInt.correcto = false;
                    respInt.mensaje = mensaje;
                    respInt.detalle = "";
                }

            }
            else
            {
                mensaje = "No hay informacion para exportar a CSV";
                Console.WriteLine(mensaje);
                respInt.correcto = false;
                respInt.mensaje = mensaje;
                respInt.detalle = "";
            }
            return respInt;
        }

        // Exporta CSVs de la carpeta local a un SFTP
        public static async Task<RespuestaInterna> ExportarFTP(string CarpetaOrigen, string CarpetaDestino, SftpConfig config, string model)
        {

            var sftpService = new SftpService(new NullLogger<SftpService>(), config);
            RespuestaInterna respInt = new RespuestaInterna();
            string mensaje = "";

            // Cargar la configuración desde appsettings.json
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory()) // Ajusta el directorio si es necesario
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            IConfiguration configApp = builder.Build();

            // Obtener los nombres de los archivos desde la configuración
            var archivosConfigurados = configApp.GetSection(model).Get<List<string>>();

            // Obtener todos los archivos .csv en la carpeta de origen
            string[] allfiles = Directory.GetFiles(CarpetaOrigen, "*.csv", System.IO.SearchOption.AllDirectories);

            // Filtrar solo los archivos que estén en la lista configurada
            var archivosFiltrados = allfiles.Where(file => archivosConfigurados.Contains(Path.GetFileName(file))).ToArray();

            foreach (string item in archivosFiltrados)
            {
                try
                {
                    var testFile = Path.Combine(CarpetaOrigen, item);

                    string[] cadenas = item.Split("\\");
                    string nombre_archivo = cadenas[cadenas.Length - 1];

                    string archivosftp = Path.Combine(CarpetaDestino, nombre_archivo).Replace("\\", "/");

                    sftpService.UploadFile(testFile, archivosftp);

                    mensaje = "ExportToSFTP: El archivo '" + item + "' fue exportado al SFTP del modelo " + model;
                    Console.WriteLine(mensaje);

                    respInt.correcto = true;
                    respInt.mensaje = "";
                    respInt.detalle = "";
                }
                catch (Exception ex)
                {
                    mensaje = "ExportToSFTP: Error - El archivo '" + item + "' no fue exportado al SFTP.\n"
                                               + ex.Message;
                    respInt.correcto = false;
                    respInt.mensaje = mensaje;
                    respInt.detalle = ex.Source;
                }

            }
            return respInt;
        }

        // Descarga CSVs del SFTP a la ruta local
        public static void DownloadCSV(string CarpetaOrigen, string CarpetaDestino, SftpConfig config, string model)
        {
            var sftpService = new SftpService(new NullLogger<SftpService>(), config);

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            IConfiguration configApp = builder.Build();

            if (!Directory.Exists(Path.GetDirectoryName(CarpetaDestino)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(CarpetaDestino));
            }

            List<string> TablasToReplicaICM_DB = configApp.GetSection(model).Get<List<string>>();

            if (TablasToReplicaICM_DB is not null)
            {
                foreach (var file in TablasToReplicaICM_DB)
                {
                    try
                    {
                        sftpService.DownloadFile(CarpetaOrigen + "|" + file, CarpetaDestino);
                        Console.WriteLine("DownloadCSV: El archivo '" + file + "' del modelo " + model + " fue descargado con exito.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("DownloadCSV: Error - El archivo '" + file + "' no se pudo descargar del modelo " + model + " del SFTP.\n" + ex.Message);
                    }
                }
            }
            else
            {
                Console.WriteLine("No se han agregado tablas para migrar a SQL");
            }
        }

        // Formatea los CSVs de la ruta local para despues ser migrados a SQL
        public static RespuestaInterna CSVToQuery(string ubiRutaCarpetaDestinoTablasToSQL, string bdcnServer, string bdcnBD, string bdcnUsuario, string bdcnContraseña)
        {
            RespuestaInterna respInt = new RespuestaInterna();
            string mensaje = "";
            List<String> files = Directory.GetFiles(ubiRutaCarpetaDestinoTablasToSQL, "*.csv").ToList();         

            if (files.Count > 0)
            {
                foreach (String file in files)
                {
                    System.Data.DataTable dt = new System.Data.DataTable();
                    string nombreTabla = file.Split("\\").Last().Split("-")[0];

                    var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        Delimiter = ";",
                        HasHeaderRecord = true,
                    };
                    try
                    {
                        var reader = new StreamReader(file);
                        var archivo = new CsvReader(reader, config);
                        var filas = archivo.GetRecords<dynamic>().ToList();

                        if (filas.Count > 0)
                        {
                            foreach (var fila in filas)
                            {                                
                                var dictionaryRow = fila as IDictionary<string, object>;
                                if (dictionaryRow != null)
                                {
                                    foreach (var column in dictionaryRow)
                                    {
                                        //If para cambiar el formato de varicent dd/MM/yyyy a el de SQL yyyy-mm-dd
                                        if (column.Value is string colValue && DateTime.TryParseExact(colValue, "d/M/yyyy", new CultureInfo("es-ES"), DateTimeStyles.None, out DateTime fecha))
                                        {
                                            dictionaryRow[column.Key] = fecha.ToString("yyyy-MM-dd");
                                        }
                                    }
                                }
                            };
                        }
                        else
                        {
                            Console.WriteLine("La tabla no tiene ninguna Fila.");
                        }
                        reader.Close();

                        // Aplicamos lo cambios en el CSV
                        var writer = new StreamWriter(file);
                        var CsvWriter = new CsvWriter(writer, config);
                        CsvWriter.WriteRecords(filas);
                        writer.Close();

                        using (TextFieldParser parser = new TextFieldParser(file))
                        {
                            parser.TextFieldType = FieldType.Delimited;
                            parser.SetDelimiters(";");

                            // Leer la primera línea con los nombres de las columnas
                            string[] columnNames = parser.ReadFields();
                            foreach (var column in columnNames)
                            {
                                dt.Columns.Add(column);
                            }

                            // Leer las filas de datos
                            while (!parser.EndOfData)
                            {
                                string[] fields = parser.ReadFields();

                                for (int i = 0; i < fields.Length; i++)
                                {
    
                                    if (string.IsNullOrWhiteSpace(fields[i])) 
                                    {
                                        fields[i] = null; 
                                    }
                                }

                                dt.Rows.Add(fields);
                            }
                        }

                        bool resInsert = false;
                        string tblInsert = "";

                        if (nombreTabla.Contains("Manufactura"))
                        {
                            resInsert = ExecuteQuery(dt, "ReplicaICM_MNF", bdcnServer, bdcnBD, bdcnUsuario, bdcnContraseña);
                            tblInsert = "ReplicaICM_MNF";
                        }
                        else if (nombreTabla.Contains("Quincenal"))
                        {
                            resInsert = ExecuteQuery(dt, "ReplicaICM_COM", bdcnServer, bdcnBD, bdcnUsuario, bdcnContraseña);
                            tblInsert = "ReplicaICM_COM";
                        }
                        else if (nombreTabla.Contains("Catorcenal"))
                        {
                            resInsert = ExecuteQuery(dt, "ReplicaICM_CAT", bdcnServer, bdcnBD, bdcnUsuario, bdcnContraseña);
                            tblInsert = "ReplicaICM_CAT";
                        }
                        else { Console.WriteLine("El archivo "+ nombreTabla +" no contiene entre su nombre algunas de las palabras clave (Catorcenal, Quincenal o Manufactura)"); }

                        mensaje = "CsvToQuery: El archivo '" + nombreTabla + "' fue migrado con exito a la tabla "+ tblInsert +" ";

                        if (resInsert)
                        {
                            Console.WriteLine(mensaje);
                            respInt.correcto = true;
                            respInt.mensaje = "";
                            respInt.detalle = "";
                        }
                        else
                        {
                            mensaje = "CsvToQuery Error: El archivo '" + nombreTabla + "' no pudo ser migrado debido a algun posible error en la insercion.";
                            Console.WriteLine(mensaje);
                            respInt.correcto = false;
                            respInt.mensaje = mensaje;
                        }
                    }
                    catch (Exception ex)
                    {
                        mensaje = "CsvToQuery Error: El archivo '" + nombreTabla + "' no pudo ser migrado debido a algun posible error." + ex.Message;
                        respInt.correcto = false;
                        respInt.mensaje = mensaje;
                        respInt.detalle = ex.Source;
                    }
                }

            }
            else
            {
                mensaje = "No existen archivos en el directorio que procesar.\n";
                respInt.correcto = true;
                respInt.mensaje = mensaje;
            }
            return respInt;
        }     

        // Insert hacia SQL en base a un DataTable
        public static bool ExecuteQuery(System.Data.DataTable dt, string tabla, string bdcnServer, string bdcnBD, string bdcnUsuario, string bdcnContraseña)
        {
            string connString = "Server="+bdcnServer+";Database="+bdcnBD+";User Id="+bdcnUsuario+";Password="+bdcnContraseña+";";
            using (SqlConnection conn = new SqlConnection())
            {
                conn.ConnectionString = connString;
                try
                {
                    conn.Open();

                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
                    {
                        bulkCopy.DestinationTableName = tabla; 
                        bulkCopy.WriteToServer(dt); 
                    }
                    return true;
                }
                catch(Exception ex)
                {
                    return false;
                }
            }        
        }
    }
}
