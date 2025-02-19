using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelosExportacion
{
    public class ConexionBD
    {

        private SqlConnectionStringBuilder _cadenaConexion;

        public ConexionBD(string servidor, string baseDatos, string usuario, string contra)
        {

            SqlConnectionStringBuilder stringBuilder = new SqlConnectionStringBuilder();

            stringBuilder.DataSource = servidor;
            stringBuilder.InitialCatalog = baseDatos;
            stringBuilder.UserID = usuario;
            stringBuilder.Password = contra;
          
      
            this._cadenaConexion = stringBuilder;

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
                            if (reader.HasRows)
                            {
                                try
                                {
                                    tabla.Load(reader);
                                    respInt.tabla = tabla;
                                    respInt.correcto = true;
                                    respInt.mensaje = "";
                                    respInt.detalle = "";                              
                                }
                                catch (Exception ex)
                                {
                                    respInt.correcto = false;
                                    respInt.tabla = tabla;
                                    respInt.mensaje = ex.Message;
                                    respInt.detalle = ex.StackTrace;
                                                               
                                }

                               
                            }else
                            {
                                respInt.correcto = false;
                                respInt.mensaje = "No hay datos";
                                
                                respInt.detalle = "Sin Datos en la tabla";
                               
                            }    
          
                            reader.Close();
                        }

                    }
                    connection.Close();
                }

                return respInt;
            }
            catch (SqlException sqlex)
            {
                string mensaje = sqlex.Message;
                if (sqlex.Number == 208)
                {
                    mensaje = "La tabla no existe";
                }
                respInt.correcto = false;
                respInt.mensaje = mensaje;
                respInt.detalle = sqlex.StackTrace;
                return respInt;
            }


        }

        public async Task<bool> VerificaBegda(string tabla) {
            try
            {
                using (SqlConnection connection = new SqlConnection(this._cadenaConexion.ConnectionString))
                {
                    await connection.OpenAsync();
                    string script = "SELECT COUNT(*) FROM sys.columns WHERE object_id = OBJECT_ID('" + tabla + "') AND name = 'BEGDA'";

                    using (SqlCommand command = new SqlCommand(script, connection))
                    {
                        int count = (int)command.ExecuteScalar();
                        if (count > 0)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                            
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error en VerificaBegda(): " + ex.Message);
            }
            return false;
        }
    }
}
