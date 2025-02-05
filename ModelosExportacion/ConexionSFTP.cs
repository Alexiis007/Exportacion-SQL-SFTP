﻿using Microsoft.Extensions.Logging;
using Renci.SshNet;
using Renci.SshNet.Sftp;
using System.Collections.Generic;

namespace ModelosExportacion
{
    public class ConexionSFTP: ISftpService
    {

        //private string strServidor;
        //private string strUsuario;
        //private string strContraseña;
        //private int strPuerto;




        //public ConexionSFTP(string Servidor, string Usuario, string Contraseña, int Puerto)
        //{
        //    strServidor = Servidor;
        //    strUsuario = Usuario;
        //    strContraseña = Contraseña;
        //    strPuerto = Puerto;



        //}



        //public async Task<bool> probarConexion()
        //{
        //    var client = new SftpClient(strServidor, strPuerto, strUsuario, strContraseña);

        //    try
        //    {
        //        client.Connect();

        //        Console.WriteLine("Prueba Exitosa de Conexion a SFTP");

        //        //client.Disconnect();
        //        return true;
        //    }
        //    catch (Exception)
        //    {

        //        return false;
        //    }


        //}


        //public async Task<RespuestaInterna> CopiarArchivo(string RutaLocal, string RutaDestino, string NombreArchivo)
        //{
        //    RespuestaInterna respInt = new RespuestaInterna();

        //    //var localPath = @"C:\path\to\local\file.txt";
        //    //var remotePath = "/Data/TablasImportadas/file.txt";

        //    var client = new SftpClient(strServidor, strPuerto, strUsuario, strContraseña);


        //    string mensaje = "";

        //    int attempts = 0;
        //    do
        //    {
        //        client.Connect();

        //        try
        //        {

        //            if (client.IsConnected)
        //            {
        //                using (var fileStream = new FileStream(RutaLocal, FileMode.Open))
        //                {

        //                    client.UploadFile(fileStream, RutaDestino);

        //                    Console.WriteLine(mensaje);
        //                    mensaje = " Archivo: " + NombreArchivo + " fue copiado correctamente al SFTP !";
        //                    respInt.correcto = true;
        //                    respInt.mensaje = mensaje;
        //                }
        //            }

        //        }
        //        catch (Exception ex)
        //        {
        //            attempts++;
        //            Console.WriteLine(ex.Message);
        //            respInt.correcto = false;
        //            respInt.mensaje = ex.Message;
        //            respInt.detalle = ex.StackTrace;
        //        }

        //    } while (attempts < 5 && !client.IsConnected);





        //    return respInt;

        //}


        private readonly ILogger<ConexionSFTP> _logger;
        private readonly SftpConfig _config;

        public ConexionSFTP(ILogger<ConexionSFTP> logger, SftpConfig sftpConfig)
        {
            _logger = logger;
            _config = sftpConfig;
        }

        /// <summary>
        /// Get list of all the files in remote directory
        /// </summary>
        /// <param name="remoteDirectory">is the directory path.</param>
        /// <returns>IEnumerable<SftpFile> , containing collection of files in the given directory</returns>

        public IEnumerable<SftpFile> ListAllFiles(string remoteDirectory = ".")
        {
            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            try
            {
                client.Connect();
                return (IEnumerable < SftpFile >) client.ListDirectory(remoteDirectory);
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, $"Failed to list files under [{remoteDirectory}]");
                return null;
            }
            finally
            {
                client.Disconnect();
            }
        }

        /// <summary>
        /// Upload the files from a local directory to the remote directory
        /// </summary>
        /// <param name="localFilePath">is the source path of the file .</param>
        /// <param name="remoteFilePath">is the destination path of the file .</param>
        /// <returns>void</returns>
        public void UploadFile(string localFilePath, string remoteFilePath)
        {
            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            try
            {
                client.Connect();
                using var s = File.OpenRead(localFilePath);
                client.UploadFile(s, remoteFilePath);
                _logger.LogInformation($"Finished uploading file [{localFilePath}] to [{remoteFilePath}]");
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, $"Failed to upload file [{localFilePath}] to [{remoteFilePath}]");
            }
            finally
            {
                client.Disconnect();
            }
        }

        /// <summary>
        /// download the files from remote location to local path
        /// </summary>
        /// <param name="localFilePath">is the source path of the file .</param>
        /// <param name="remoteFilePath">is the destination path of the file .</param>
        /// <returns>void</returns>
        public void DownloadFile(string remoteFilePath, string localFilePath)
        {
            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            try
            {
                client.Connect();
                using var s = File.Create(localFilePath);
                client.DownloadFile(remoteFilePath, s);
                _logger.LogInformation($"Finished downloading file [{localFilePath}] from [{remoteFilePath}]");
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, $"Failed to download file [{localFilePath}] from [{remoteFilePath}]");
            }
            finally
            {
                client.Disconnect();
            }
        }

        /// <summary>
        /// delete the remote files 
        /// </summary>     
        /// <param name="remoteFilePath">is the path of the file to be deleted .</param>
        /// <returns>void</returns>
        public void DeleteFile(string remoteFilePath)
        {
            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            try
            {
                client.Connect();
                client.DeleteFile(remoteFilePath);
                _logger.LogInformation($"File [{remoteFilePath}] deleted.");
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, $"Failed to delete file [{remoteFilePath}]");
            }
            finally
            {
                client.Disconnect();
            }
        }

    }
}
