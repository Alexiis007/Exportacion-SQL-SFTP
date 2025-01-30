using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;
using ModelosExportacion;

namespace Exportacion
{
    public class Conexion
    {
        private SqlConnectionStringBuilder _cadenaConexion;

        public Conexion(SqlConnectionStringBuilder stringBuilder)
        {
            this._cadenaConexion = stringBuilder;

            //Setting TLS 1.2 protocol
            System.Net.ServicePointManager.Expect100Continue = true;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls12;

        }

        public Conexion(string servidor, string usuario, string contraseña, string basededatos)
        {
            SqlConnectionStringBuilder stringBuilder = new SqlConnectionStringBuilder();

            stringBuilder.DataSource = servidor;
            stringBuilder.UserID = usuario;
            stringBuilder.Password = contraseña;
            stringBuilder.InitialCatalog = basededatos;

            this._cadenaConexion = stringBuilder;

            //Setting TLS 1.2 protocol
            System.Net.ServicePointManager.Expect100Continue = true;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls12;

        }

        public async Task<bool> probarConexion()
        {
            SqlConnection cnn = new SqlConnection(this._cadenaConexion.ConnectionString);

            try
            {
                await cnn.OpenAsync();
                cnn.Close();

                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<RespuestaInterna> ejecutScript(string script)
        {
            RespuestaInterna respInt = new RespuestaInterna();
            DataTable tabla = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(this._cadenaConexion.ConnectionString))
                {
                    await connection.OpenAsync();

                    String sql = script;

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                try
                                {
                                    //Usuarios musuario = new Usuarios();
                                    //musuario.usuario = reader.GetInt32(0).ToString();
                                    //musuario.contraseña = reader.GetString(1).ToString();

                                    //lusuarios.Add(musuario);
                                }
                                catch (Exception ex)
                                {
                                    respInt.correcto = false;
                                    respInt.mensaje = ex.Message;
                                    respInt.detalle = ex.StackTrace;
                                    return respInt;
                                }
                            }
                            tabla.Load(reader);
                        }
                       
                    }
                    connection.Close();
                }

                //respInt.objeto = lusuarios;
                //respInt.horaFinal = DateTime.Now;

                return respInt;
            }
            catch (SqlException sqlex)
            {
                respInt.correcto = false;
                respInt.mensaje = sqlex.Message;
                respInt.detalle = sqlex.StackTrace;
                return respInt;
            }


        }

        public async Task<RespuestaInterna> ejecutaSP(string sp_nombre, SqlParameter[] parametros)
        {
            RespuestaInterna resp = new RespuestaInterna();

            try
            {
                List<string> nombres = new List<string>();

                using (SqlConnection cn = new SqlConnection(this._cadenaConexion.ConnectionString))
                {
                    cn.Open();

                    SqlCommand cmd = new SqlCommand(sp_nombre, cn);
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.AddRange(parametros);

                    SqlDataReader dr = await cmd.ExecuteReaderAsync();

                    if (dr.Read())
                    {
                        for (int i = 0; i < dr.FieldCount; i++)
                        {
                            string nombre = dr.GetString(i);
                            nombres.Add(nombre);
                        }

                    }

                    cn.Close();
                }

                resp = generaRespuesta(resp, nombres, this.GetType().FullName);

                return resp;
            }
            catch (Exception ex)
            {
                return generaError(resp, ex.Message, ex.StackTrace);
            }
        }

        public async Task<RespuestaInterna> execSP(string sp_nombre, SqlParameter[] parametros = null)
        {
            RespuestaInterna resp = new RespuestaInterna();
            DataTable tabla = new DataTable();

            try
            {
                List<string> nombres = new List<string>();

                int cnt = 1;
                while (!await probarConexion())
                {
                    if (cnt > 5)
                        throw new Exception("No se pudo establecer conexión con la base de datos.");

                    cnt++;

                    await Task.Delay(TimeSpan.FromMinutes(5));
                }

                using (SqlConnection cn = new SqlConnection(this._cadenaConexion.ConnectionString))
                {

                    cn.Open();

                    SqlCommand cmd = new SqlCommand(sp_nombre, cn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    if (parametros != null)
                        cmd.Parameters.AddRange(parametros);

                    var dr = await cmd.ExecuteReaderAsync();

                    tabla.Load(dr);

                    cn.Close();
                }

                resp = generaRespuesta(resp, tabla, this.GetType().FullName);

                return resp;
            }
            catch (Exception ex)
            {
                return generaError(resp, ex.Message, ex.StackTrace);
            }
        }


        private RespuestaInterna generaError(RespuestaInterna resp, string mensaje = "Error en el proceso", string detalle = "El proceso no se llevó a cabo de la mera correcta.")
        {
            resp.correcto = false;
            resp.objeto = null;
            resp.mensaje = mensaje;
            resp.detalle = detalle;
            resp.horaFinal = DateTime.Now;

            return resp;
        }
        private RespuestaInterna generaRespuesta(RespuestaInterna resp, object objeto, string funcion, string detalle = "El proceso se terminó de manera correcta.")
        {
            resp.objeto = objeto;
            resp.detalle = "Se ejecutó: " + funcion + Environment.NewLine + detalle;
            resp.horaFinal = DateTime.Now;

            return resp;
        }
    }
}
