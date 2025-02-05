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

namespace ModelosExportacion
{
    static class Funciones
    {
        public static bool ExportarExcel(System.Data.DataTable tabla,  string NombreTabla, string RutaDestino)
        {
            if (tabla == null || tabla.Columns.Count == 0)
            {
                return false;
            }

              

            var excelApp = new Excel.Application();
            excelApp.Workbooks.Add();

            // single worksheet
            Microsoft.Office.Interop.Excel._Worksheet Worksheet = (Microsoft.Office.Interop.Excel._Worksheet)excelApp.ActiveSheet;


            // column headings
            for (var i = 0; i < tabla.Columns.Count; i++)
            {
                Worksheet.Cells[1, i + 1] = tabla.Columns[i].ColumnName;
            }

            // rows
            for (var i = 0; i < tabla.Rows.Count; i++)
            {
                // to do: format datetime values before printing
                for (var j = 0; j < tabla.Columns.Count; j++)
                {
                    Worksheet.Cells[i + 2, j + 1] = tabla.Rows[i][j];
                }
            }

            // check file path
            if (!string.IsNullOrEmpty(RutaDestino))
            {
                try
                {
                    Worksheet.SaveAs(RutaDestino);
                    excelApp.Quit();
                    Console.WriteLine("ExportToExcel: Tabla " + NombreTabla + " exportada a Excel !");
                }
                catch (Exception ex)
                {
                    throw new Exception("ExportToExcel: Error - La tabla " + NombreTabla + " no fue exportada Excel verifica la Ruta.\n"
                                        + ex.Message);
                }
            }
            else
            { // no file path is given
                excelApp.Visible = true;
            }



            return true;

        }

   
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

                            //Version donato con replace, con ";"  sale desacomodada e incompleta. Con "," sale acomodada pero incompleta
                            List<string> valoresxlinea = tabla.AsEnumerable().Select(row => string.Join(",", row.ItemArray)).ToList();
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

                        foreach (var file in ArchivosCSV) {
                            if (file != ruta + "\\Heineken.Sociedades.csv" && file != ruta + "\\Heineken.DivisionesPersonal.csv") {

                                //Cargamos el archivo CSV con el reader (archivo de lectura)
                                var reader = new StreamReader(file);

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

                                        }
                                    }
                                }

                                reader.Close();

                                //----- Guardamos el archivo csv -----
                                // abrimos el archivo csv en escritura
                                var writer = new StreamWriter(file);
                                // convertimos ese archivo de escritura a un archivo csv de escritura
                                var CsvWriter = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture));
                                //Ahora modificamos el archivo csv pasandole todas las filas modificadas
                                CsvWriter.WriteRecords(filas);
                                writer.Close();
                            }
                        }

                            mensaje = "ExportToCSV: Tabla '" + NombreTabla + "' exportada a csv con " + tabla.Rows.Count.ToString() + " filas";
                        Console.WriteLine(mensaje);

                        respInt.correcto = true;
                        respInt.mensaje = mensaje;
                        respInt.detalle = "";
                    }
                    catch (Exception ex)
                    {

                        mensaje = "ExportToCSV: Error - La tabla '" + NombreTabla + "'.\n"
                                               + ex.Message;
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

        public static RespuestaInterna ExportarExcelEppPlus(System.Data.DataTable tabla, string NombreTabla, string RutaDestino)
        {
            string mensaje = "";

            RespuestaInterna respInt = new RespuestaInterna();

            if (tabla != null && tabla.Rows.Count > 0)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add(NombreTabla);

                    // Encabezados
                    for (var i = 0; i < tabla.Columns.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = tabla.Columns[i].ColumnName;
                    }

                    int x = 1;
                    // Filas
                    for (var i = 0; i < tabla.Rows.Count; i++)
                    {      
                        for (var j = 0; j < tabla.Columns.Count; j++)
                        {
                            ws.Cells[x + 1, j + 1].Value =  tabla.Rows[i][j];
                        }
                        x++;
                    }

                    if (!string.IsNullOrEmpty(RutaDestino))
                    {
                        try
                        {
                           
                            var file = new FileInfo(RutaDestino);
                            package.SaveAs(file);

                            mensaje = "ExportToExcel: Tabla '" + NombreTabla + "' exportada a Excel con " + tabla.Rows.Count.ToString() + " filas";
                            Console.WriteLine(mensaje);

                            respInt.correcto = true;
                            respInt.mensaje = mensaje;
                            respInt.detalle = "";
                            

                        }
                        catch (Exception ex)
                        {
                            mensaje = "ExportToExcel: Error - La tabla '" + NombreTabla + "' no fue exportada Excel verifica la Ruta.\n"
                                                + ex.Message;
                            Console.WriteLine(mensaje);

                            respInt.correcto = false;
                            respInt.mensaje = mensaje;
                            respInt.detalle = ex.StackTrace;


                        }
                    }
                    else
                    { // no file path is given
                        mensaje = "ExportToExcel: La ruta " + RutaDestino + " No es valida!";
                        Console.WriteLine(mensaje);

                        respInt.correcto = false;
                        respInt.mensaje = mensaje;
                        respInt.detalle = "";
                    }

                  

                }
            }
            else
            {
                mensaje = "No hay informacion para exportar a Excel";
                Console.WriteLine(mensaje);
                respInt.correcto = false;
                respInt.mensaje = mensaje;
                respInt.detalle = "";
            }

             return respInt;

        }

        public static RespuestaInterna ExportarSFTP(string CarpetaOrigen ,string CarpetaDestino, SftpConfig config)
        {

            var sftpService = new SftpService(new NullLogger<SftpService>(), config);
            RespuestaInterna respInt = new RespuestaInterna();
            string mensaje = "";

            string[] files = Directory.GetFiles(CarpetaOrigen, "*.csv", SearchOption.AllDirectories);
            foreach (string item in files)
            {
                try
                {
                    var testFile = Path.Combine(CarpetaOrigen, item);
                    //var testFile = Path.Combine("C:\\Destino", item);

                    string[] cadenas =  item.Split("\\");
                    string nombre_archivo =  cadenas[cadenas.Length - 1];

                    //string archivosftp = Path.Combine(CarpetaDestino, nombre_archivo);
                    string archivosftp  = Path.Combine("//Data//" , nombre_archivo);

                    sftpService.UploadFile(testFile, archivosftp);

                    mensaje = "ExportToSFTP: El archivo '" + item + "' fue exportado al SFTP";
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

        // Nuevo Metodo
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



    }

}
