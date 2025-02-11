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


namespace ModelosExportacion
{
    static class Funciones
    {
        // Exporta tablas SQL a CSVs y las guarda en carpeta local
        public async static Task<RespuestaInterna> ExportCSV(System.Data.DataTable tabla, string NombreTabla, string RutaDestino)
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
                                    // Rodear con comillas dobles y reemplazar las comillas dobles internas por comillas dobles duplicadas
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
                        string ruta = @"C:\Destino";
                        //Nombres archivos CSV ubicados en la ruta
                        string[] ArchivosCSV = Directory.GetFiles(ruta, "*.csv");

                        //Cargamos el archivo CSV con el reader (archivo de lectura)
                        var reader = new StreamReader(RutaDestino);

                        //Cargamos el archivo reader como como un archivo de tipo csv de solo lectura
                        var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));

                        //Obtenemos la filas con un tipo de valor de retorno dynamic
                        //Y ese retorno lo pasamos a una lista con la funcion ToList()
                        var filas = csv.GetRecords<dynamic>().ToList();

                        //Recoremos la lista de filas del csv
                        foreach (var fila in filas)
                        {
                            var diccionarioFila = fila as IDictionary<string, object>;
                            //Recorremos las columnas de esta fila
                            foreach (var col in diccionarioFila)
                            {
                                //Verificamos el tipo de dato de la col
                                // Si es de tipo DateTime lo pasamos a Date
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

                        reader.Close();

                        //----- Guardamos el archivo csv -----
                        // abrimos el archivo csv en escritura
                        var writer = new StreamWriter(RutaDestino);
                        // convertimos ese archivo de escritura a un archivo csv de escritura
                        var CsvWriter = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture));
                        //Ahora modificamos el archivo csv pasandole todas las filas modificadas
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

        // Exporta CSVs de la carpeta local al SFTP
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
            string[] allfiles = Directory.GetFiles(CarpetaOrigen, "*.csv", SearchOption.AllDirectories);

            // Filtrar solo los archivos que estén en la lista configurada
            var archivosFiltrados = allfiles.Where(file => archivosConfigurados.Contains(Path.GetFileName(file))).ToArray();

            foreach (string item in archivosFiltrados)
            {
                try
                {
                    var testFile = Path.Combine(CarpetaOrigen, item);
                    //var testFile = Path.Combine("C:\\Destino", item);

                    string[] cadenas = item.Split("\\");
                    string nombre_archivo = cadenas[cadenas.Length - 1];

                    //string archivosftp = Path.Combine(CarpetaDestino, nombre_archivo);
                    string archivosftp = Path.Combine("//Data//", nombre_archivo);

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

        // Descarga CSVs del SFTP a carpeta local
        public static async Task<RespuestaInterna> DownloadCSV(string CarpetaOrigen, string CarpetaDestino, SftpConfig config, string model)
        {

            var sftpService = new SftpService(new NullLogger<SftpService>(), config);

            RespuestaInterna respInt = new RespuestaInterna();
            string mensaje = "";

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

                        sftpService.DownloadFile(CarpetaOrigen+"|"+ file, CarpetaDestino);

                        mensaje = "DownloadCSV: El archivo '" + file + "' del modelo " + model + " fue descargado con exito.";

                        Console.WriteLine(mensaje);
                        respInt.correcto = true;
                        respInt.mensaje = "";
                        respInt.detalle = "";
                    }
                    catch (Exception ex)
                    {
                        mensaje = "DownloadCSV: Error - El archivo '" + file + "' no se pudo descargar del modelo " + model + " del SFTP.\n" + ex.Message;
                        respInt.correcto = false;
                        respInt.mensaje = mensaje;
                        respInt.detalle = ex.Source;

                    }
                }
            }
            else
            {
                Console.WriteLine("No se han agregado tablas para migrar a SQL");
            }

            return respInt;

        }

        public static RespuestaInterna CSVToQuery(string ubiRutaCarpetaDestinoTablasToSQL, string bdcnServer, string bdcnBD, string bdcnUsuario, string bdcnContraseña)
        {
            RespuestaInterna respInt = new RespuestaInterna();
            string mensaje = "";

            List<String> files = Directory.GetFiles(ubiRutaCarpetaDestinoTablasToSQL, "*.csv").ToList();
            if (files.Count > 0)
            {
                foreach (String file in files)
                {
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

                        string tabla = file.Replace(".csv", "");
                        List<string> AllDataInsert = new List<string>();

                        if (filas.Count > 0)
                        {
                            // Convertir la fila a IDictionary<string, object> para poder acceder a las claves
                            var firstRow = filas[0] as IDictionary<string, object>;
                            if (firstRow != null)
                            {
                                // Recorrer las claves del diccionario (nombres de las columnas)
                                string cols = "";
                                foreach (var column in firstRow.Keys)
                                {
                                    cols += column + ",";
                                }
                                AllDataInsert.Add("INSERT INTO " + "ReplicaICM_MNF" + "(" + cols.Substring(0, cols.Length - 1) + ") VALUES");
                            }
                            foreach (var fila in filas)
                            {
                                var dictionaryRow = fila as IDictionary<string, object>;
                                if (dictionaryRow != null)
                                {
                                    string insertValues = "(";
                                    foreach (var column in dictionaryRow)
                                    {
                                        //If para cambiar el formato de varicent dd/MM/yyyy a el de SQL yyyy-mm-dd
                                        if (column.Value is string colValue && DateTime.TryParseExact(colValue, "d/M/yyyy", new CultureInfo("es-ES"), DateTimeStyles.None, out DateTime fecha))
                                        {
                                            insertValues += "\'" + fecha.ToString("yyyy-MM-dd") + "\',";
                                        }
                                        else
                                        {
                                            insertValues += "\'" + column.Value + "\',";
                                        }
                                    }
                                    insertValues = insertValues.Substring(0, insertValues.Length - 1);
                                    insertValues += "),";
                                    AllDataInsert.Add(insertValues);
                                }
                            }
                            string ultimaCadena = AllDataInsert[AllDataInsert.Count - 1];
                            AllDataInsert[AllDataInsert.Count - 1] = AllDataInsert[AllDataInsert.Count - 1].Substring(0, ultimaCadena.Length - 1) + ';';
                        }
                        else
                        {
                            Console.WriteLine("La tabla no tiene ninguna Fila.");
                        }
                        foreach (var a in AllDataInsert)
                        {
                            Console.WriteLine(a);
                        }

                        reader.Close();

                        mensaje = "Procesamiento CsvToQuery: El archivo '" + file + "' fue convertido a Query con exito.";
                        Console.WriteLine(mensaje);
                        respInt.correcto = true;
                        respInt.mensaje = "";
                        respInt.detalle = "";
                    }
                    catch (Exception ex)
                    {
                        mensaje = "Procesamiento CsvToQuery: Error - El archivo '" + file + "' no pudo ser convertido a Query.\n" + ex.Message;
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

        public static void ExecuteQuery(string query, string bdcnServer, string bdcnBD, string bdcnUsuario, string bdcnContraseña)
        {
            try
            {
                string connString = "Server="+bdcnServer+";Database="+bdcnBD+";User Id="+bdcnUsuario+";Password="+bdcnContraseña+";";
                using (SqlConnection conn = new SqlConnection())
                {
                    SqlCommand command = new SqlCommand(query, conn);

                    try
                    {
                        conn.Open();

                        int rowsAffected = command.ExecuteNonQuery();
                        Console.WriteLine($"Filas insertadas: {rowsAffected}");
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine("Error"+ ex.Message);
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error:",ex.Message);
            }
        }
    }
}
