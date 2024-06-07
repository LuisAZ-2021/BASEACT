using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Base
{
    public class Usuario
    {
        private int id;
        private string name;
        private string password;
        private int rol;

        public int ID { get; set; }

        public string Nombre { get; set; }

        public string Password { get; set; }

        public int Rol { get; set; }

        public Usuario()
        {

        }

        public Usuario(int id, string nombre, string password, int rol)
        {
            this.id = id;
            this.name = nombre;
            this.password = password;
            this.rol = rol;
        }

        public void Crear()
        {

        }

        public void Guardar(string datosDetalles)
        {
            string guardo = "Guardo Datos en Localidad";
            string detalles = "Se guardó: " + datosDetalles;

            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();

                    string query = "INSERT INTO Bitacora(Fecha, IdUsuario, Acción, Descripción)" +
                        "VALUES(GETDATE(), @idusuario, @acción, @detalle)";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    cmd.Parameters.AddWithValue("idusuario", this.ID);
                    cmd.Parameters.AddWithValue("acción", guardo);
                    cmd.Parameters.AddWithValue("detalle", detalles);

                    cmd.ExecuteNonQuery();

                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al Registrar Movimiento:" + ex.Message);
            }

        }

        public void Modificar()
        {

        }

        public void Eliminar()
        {

        }
    }
}
