using Microsoft.Extensions.Logging;
using Renci.SshNet;
using Renci.SshNet.Sftp;
using System.Net.Http.Json;
using System.Runtime.InteropServices.JavaScript;

namespace SFTPService
{
    /// <summary>
    /// Class containing sftp file operations
    /// </summary>
    public class SftpService : ISftpService
    {
        private readonly ILogger<SftpService> _logger;
        private readonly SftpConfig _config;

        public SftpService(ILogger<SftpService> logger, SftpConfig sftpConfig)
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
                return (IEnumerable<SftpFile>)client.ListDirectory(remoteDirectory);
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
        public bool UploadFile(string localFilePath, string remoteFilePath)
        {
            bool ArchivoCargado = false;   
            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            try
            {
                client.Connect();
                using var s = File.OpenRead(localFilePath);
                client.UploadFile(s, remoteFilePath);
                _logger.LogInformation($"Finished uploading file [{localFilePath}] to [{remoteFilePath}]");
                ArchivoCargado = true;
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, $"Failed to upload file [{localFilePath}] to [{remoteFilePath}]");               
            }
            finally
            {
                client.Disconnect();
            }
            return ArchivoCargado;
        }

        /// <summary>
        /// download the files from remote location to local path
        /// </summary>
        /// <param name="localFilePath">is the source path of the file .</param>
        /// <param name="remoteFilePath">is the destination path of the file .</param>
        /// <returns>void</returns>
        public void DownloadFile(string remoteFilePath, string localFilePath)
        {
            var fileName= remoteFilePath.Split("|")[1];
            remoteFilePath = remoteFilePath.Split("|")[0];

            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            try
            {
                client.Connect();

                // En base a el nombre del archivo y la fecha de modificacion sacamos el mas reciente
                var lastFile = client.ListDirectory(remoteFilePath)
                  .Where(file => !file.IsDirectory && file.Name.Contains(fileName)) 
                  .OrderByDescending(file => file.LastWriteTime) 
                  .ToList().First();

                // Si el archivo es tipo txt se pasa a csv
                if (lastFile.FullName.Contains(".txt"))
                {
                    using var s = File.Create(localFilePath + "\\" + lastFile.Name.Replace(".txt", ".csv"));
                    client.DownloadFile(remoteFilePath.Replace("\\", "") + "/" + lastFile.Name, s);
                }
                else
                {
                    using var s = File.Create(localFilePath + "\\" + lastFile);
                    client.DownloadFile(remoteFilePath.Replace("\\", "") + "/" + lastFile.Name, s);
                    client.DownloadFile(remoteFilePath.Replace("\\", "") + "/" + lastFile.Name, s);
                }
            }
            catch (Exception exception)
            {
                _logger.LogError(exception, $"Failed to download file [{localFilePath}] from [{remoteFilePath}]");
                Console.WriteLine($"Failed to download file [{localFilePath}] from [{remoteFilePath}] Error:" + exception.Message);
            }
            finally
            {
                client.Disconnect();
            }
        }

        public List<String> GetFiles(string path, SftpClient client)
        {
            using (client)
            {
                client.Connect();
                return client
                    .ListDirectory(path)
                    .Where(f => !f.IsDirectory)
                    .Select(f => f.Name)
                    .ToList();
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

        public bool UploadDirectory(string localePath, string remotePath)
        {
            bool DirectorioCargado = false;
            using var client = new SftpClient(_config.Host, _config.Port == 0 ? 22 : _config.Port, _config.UserName, _config.Password);
            //try
            //{
            //    client.Connect();
              
            //    string[] files = Directory.GetFiles(CarpetaOrigen, "*.csv", SearchOption.AllDirectories);
            //    foreach (string item in files)
            //    {
            //        if (client.IsConnected)
            //        {
                        
            //        }
            //    }
               
            //    using var s = File.OpenRead(localFilePath);
            //    client.UploadFile(s, remoteFilePath);
            //    _logger.LogInformation($"Finished uploading file [{localFilePath}] to [{remoteFilePath}]");
            //    DirectorioCargado = true;
            //}
            //catch (Exception exception)
            //{
            //    _logger.LogError(exception, $"Failed to upload file [{localFilePath}] to [{remoteFilePath}]");
            //}
            //finally
            //{
            //    client.Disconnect();
            //}

            return DirectorioCargado;
        }
    }
}
