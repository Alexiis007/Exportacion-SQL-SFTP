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


namespace Exportacion
{
    enum TipoProceso
    {
        Automatico =1,
        Manual = 2
    }
    class Program
    {

        static async Task Main(string[] args)
        {
       
          
            Configuraciones _config = new Configuraciones();
            configuracionesAppSettings config = new configuracionesAppSettings();
           

            //  # # # # # # #   VARIABLES DE CONFIGURACIÓN   # # # # # # #  //

            //Informacion BD
            config.bdcnServer = _config.obtenerParametros("ConnectionString", "bdcnServer");
            config.bdcnBD = _config.obtenerParametros("ConnectionString", "bdcnBD");
            config.bdcnUsuario = _config.obtenerParametros("ConnectionString", "bdcnUsuario");
            config.bdcnContraseña = _config.obtenerParametros("ConnectionString", "bdcnContraseña");

            //ICMMNFHeinekenQA_SFTP
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

            //Informacion Tablas BD
            config.strTablasExportar = _config.obtenerParametros("Informacion", "Tablas");
            config.intMaximoRegistros = int.Parse(_config.obtenerParametros("Informacion", "MaximoRegistros"));
  
            //Fecha Filtro Begda
            config.dteFechaFiltro = _config.obtenerParametros("Informacion", "FechaFiltro");

            //Ubicaciones de descarga y subida
            config.ubiRutaCarpetaLocalExcel = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaLocalExcel");
            config.ubiRutaCarpetaDestinoExcel = _config.obtenerParametros("Ubicaciones", "ubiRutaCarpetaDestinoExcel");

            ExportarInicio exp = new ExportarInicio(config);

            await exp.ExportarArchivosAsync();
                

        }


        //private static void ShowMenu()
        //{
        //    Console.Clear();
        //    Console.ResetColor();
        //    Console.SetCursorPosition(0, 0);
        //    Console.WriteLine("=============== Exportaciones ================");
        //    Console.WriteLine("*                                            *");
        //    Console.WriteLine("*              Menu Principal                *");
        //    Console.WriteLine("*                                            *");
        //    Console.WriteLine("==============================================");
        //    Console.ResetColor();
        //    Console.WriteLine();
        //    Console.WriteLine("{0,-10} Para Salir del programa.", "[ESC]");
        //    Console.WriteLine();         
        //    Console.WriteLine("{0,-10} Para iniciar el programa Manual.", "[M]");
        //    Console.WriteLine();
        //}


       
        //private async static void EjecutarImportacion(TipoProceso tipoProceso, ExportarInicio exp, bool EntroInicio)
        //{ 
        //    switch (tipoProceso)
        //    {
        //        case TipoProceso.Automatico:    

        //            if (!EntroInicio)
        //            {
        //                watch.Stop();
        //                watch.Restart();
        //            }
        //            await exp.ExportarArchivosAsync();

        //            Task.Delay(5000).Wait();

        //            ShowMenu();
                    

        //            watch.Start();

        //            CancellationTokenSource cts_timer_auto = new CancellationTokenSource();
        //            var task_auto = new Task(() => ShowTheWatch(cts_timer_auto, exp));
        //            task_auto.Start();

        //            if (!EntroInicio)
        //            {
        //                watch.Start();
        //            }
                    

        //            break;
        //        case TipoProceso.Manual:

        //            //Console.Clear();
        //            Console.ForegroundColor = ConsoleColor.Green;

        //            watch.Stop();
        //            watch.Restart();

        //            await exp.ExportarArchivosAsync();
                    
        //            ShowMenu();
                    
        //            watch.Start();

                   

        //            CancellationTokenSource cts_timer_manual = new CancellationTokenSource();
        //            var task_manual = new Task(() => ShowTheWatch(cts_timer_manual, exp));
        //            task_manual.Start();


        //            break;
                
        //    }
        //}

        //private static void Inicial(CancellationTokenSource _cts)
        //{
        //    while (!_cts.IsCancellationRequested)
        //    {

        //        Task.Delay(1).Wait();
        //    }
        //}
        //static void ShowTheWatch(CancellationTokenSource _cts, ExportarInicio exp)
        //{
        //    int minuteInOneHour = 60;
        //    int secondInOneMinute = 60;
        //    int milisecondInOneSecond = 1000;

        //    Task.Delay(1).Wait();
        //    while (!_cts.IsCancellationRequested)
        //    {
        //        isTimerRuning = true;
              
        //        //if (isTimerRuning)
        //        //{
        //            //Console.SetCursorPosition(60, 0);
        //            if (watch.ElapsedMilliseconds != 0)
        //            {
        //                var minute = (watch.ElapsedMilliseconds / (secondInOneMinute * milisecondInOneSecond)) % minuteInOneHour;
        //                var sec = (watch.ElapsedMilliseconds / milisecondInOneSecond) % secondInOneMinute;
        //                var miliSec = watch.ElapsedMilliseconds % milisecondInOneSecond;

        //                //Console.WriteLine("{0,2:0#}:{1,2:0#}:{2,-100:0##}", minute, sec, miliSec);

        //                if (minute == tiempo )
        //                {
        //                    isTimerRuning = false;
        //                    _cts.Cancel();

        //                    EjecutarImportacion(TipoProceso.Automatico, exp, false);

        //                }
        //            }
        //        //}
        //        Task.Delay(1).Wait();
        //    }
        //}



        //public static IHostBuilder CreateHostBuilder(string[] args) =>
        //       Host.CreateDefaultBuilder(args)
        //           .ConfigureServices((hostContext, services) =>
        //           {
        //               services.AddHostedService<Worker>();
        //   });  
    
    }
}


