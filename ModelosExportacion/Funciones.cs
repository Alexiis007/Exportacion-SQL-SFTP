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

     

        //public static void ExportTableDataToExcel(System.Data.DataTable table, string NombreTabla, string RutaDestino)
        //{
        //    RespuestaInterna respInt = new RespuestaInterna();

        //    if (table != null && table.Rows.Count > 0)
        //    {
        //        object misValue = System.Reflection.Missing.Value;
        //        Application application = new Application();
        //        application.Visible = false;
        //        Workbook workbook = application.Workbooks.Add(misValue);
        //        Worksheet worksheet = (Worksheet)workbook.Worksheets.get_Item(1);

        //        worksheet.Name = "CSharpeDataTableExportedToExcel";
        //        worksheet.Cells.Font.Size = 12;
        //        AddColumnInSheet(worksheet, table);

        //        worksheet.Activate();

        //        for (int j = 1; j <= table.Rows.Count; j++)
        //        {
        //            for (int i = 1; i <=table.Columns.Count; i++)
        //            {
        //                string data = table.Rows[j- 1].ItemArray[i - 1].ToString();
        //                worksheet.Cells[j + 1,i] = data;
        //            }
        //        }

        //        workbook.SaveAs(RutaDestino,
        //        Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
        //        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue,
        //        misValue);
        //        workbook.Close(true, misValue, misValue);
        //        application.Quit();


        //    }
        //    else
        //    {
        //        respInt.correcto = false;
        //        respInt.mensaje = "No hay informacion para exportar a Excel";
        //        respInt.detalle = "La tabla " + NombreTabla + "' no tiene filas para expotar";
        //    }



        //}

        //public static void AddColumnInSheet(Worksheet worksheet, System.Data.DataTable TempDT)
        //{
        //    for (int i = 1; i <= TempDT.Columns.Count; i++)
        //    {
        //        string data = TempDT.Columns[i - 1].ToString();
        //        worksheet.Cells[1, i] = data;
        //    }
        //}


        //public static RespuestaInterna  ExportIron(System.Data.DataTable tabla, string NombreTabla, string RutaDestino, string licenciaIron)
        //{
        //    IronXL.License.LicenseKey = licenciaIron;
        //    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
        //    WorkSheet ws = wb.DefaultWorkSheet;

        //    string mensaje = "";

        //    RespuestaInterna respInt = new RespuestaInterna();

        //    if (tabla != null && tabla.Rows.Count > 0)     
        //    {


        //        // column headings
        //        for (var i = 0; i < tabla.Columns.Count; i++)
        //        {
        //            ws.SetCellValue(0, i, tabla.Columns[i].ColumnName);
        //        }

        //        // rows
        //        for (var i = 0; i < tabla.Rows.Count; i++)
        //        {
        //            // to do: format datetime values before printing
        //            for (var j = 0; j < tabla.Columns.Count; j++)
        //            {
        //                ws.SetCellValue(i + 1, j + 1, tabla.Rows[i][j]);
        //            }
        //        }

        //        // check file path
        //        if (!string.IsNullOrEmpty(RutaDestino))
        //        {
        //            try
        //            {
        //                wb.SaveAs(RutaDestino);
        //                mensaje = "ExportToExcel: Tabla '" + NombreTabla + "' exportada a Excel con " + tabla.Rows.Count.ToString() + " filas";
        //                Console.WriteLine(mensaje);

        //                respInt.correcto = true;
        //                respInt.mensaje = mensaje;
        //                respInt.detalle = "";

        //            }
        //            catch (Exception ex)
        //            {
        //                mensaje = "ExportToExcel: Error - La tabla '" + NombreTabla + "' no fue exportada Excel verifica la Ruta.\n"
        //                                    + ex.Message;
        //                Console.WriteLine(mensaje);

        //                respInt.correcto = false;
        //                respInt.mensaje = mensaje;
        //                respInt.detalle = ex.StackTrace;


        //            }
        //        }
        //        else
        //        { // no file path is given
        //            mensaje = "ExportToExcel: La ruta " + RutaDestino + " No es valida!";
        //            Console.WriteLine(mensaje);

        //            respInt.correcto = false;
        //            respInt.mensaje = mensaje;
        //            respInt.detalle = "";
        //        }


        //    }else
        //    {
        //        mensaje = "No hay informacion para exportar a Excel";
        //        Console.WriteLine(mensaje);
        //        respInt.correcto = false;
        //        respInt.mensaje = mensaje;
        //        respInt.detalle = "";
        //    }

        //    return respInt;
        //}

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
                        using (StreamWriter sw = new StreamWriter(RutaDestino))
                        {
                            StringBuilder sb = new StringBuilder();
                            string[] NombresColumnas = tabla.Columns.Cast<DataColumn>().Select(columna => columna.ColumnName).ToArray();
                            string encabezados = string.Join(",", NombresColumnas);
                            sb.AppendLine(encabezados);

                            List<string> valoresxlinea = tabla.AsEnumerable().Select(row => string.Join(",", row.ItemArray)).ToList();
                            string rows = string.Join(Environment.NewLine, valoresxlinea);
                            sb.AppendLine(rows);

                            await sw.WriteLineAsync(sb);

                            sb.Clear();
                        }

                        // ----- Configuraciones CSV ------

                        //Nombres archivos CSV ubicados en la ruta
                        string[] ArchivosCSV = Directory.GetFiles(RutaDestino, "*.csv");

                        foreach (var file in ArchivosCSV) {
                            //Cargamos el archivo CSV con el reader (archivo de lectura)
                            var reader = new StreamReader(file);

                            //Cargamos el archivo reader como como un archivo de tipo csv de solo lectura
                            var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));

                            //Obtenemos la filas con un tipo de valor de retorno dynamic
                            //Y ese retorno lo pasamos a una lista con la funcion ToList()
                            var filas = csv.GetRecords<dynamic>().ToList();

                            //Recoremos la lista de filas del csv
                            foreach (var fila in filas){
                                //Recorremos las columnas de esta fila
                                foreach (var col in fila) {
                                    //Verificamos el tipo de dato de la col
                                    // Si es de tipo DateTime lo pasamos a Date
                                    if (col.Value is DateTime) {
                                        //Cambiamos el tipo de dato de la col. en la fila 
                                        fila[col.Key] = ((DateTime)col.Value).Date.ToString("dd-mm-yyyy");
                                    }
                                }
                            }

                            //----- Guardamos el archivo csv -----
                            // abrimos el archivo csv en escritura
                            var writer = new StreamWriter(file);
                            // convertimos ese archivo de escritura a un archivo csv de escritura
                            var CsvWriter = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture));
                            //Ahora modificamos el archivo csv pasandole todas las filas modificadas
                            CsvWriter.WriteRecords(filas);
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
                { // no file path is given
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
