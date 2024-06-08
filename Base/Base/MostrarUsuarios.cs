using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Base
{
    public partial class MostrarUsuarios : Form
    {
        public MostrarUsuarios()
        {
            InitializeComponent();
        }

        public void mosUsuarios()
        {
            //clickDerecho = false;
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();
                    DataTable dt = new DataTable();
                    string query = @"SELECT l.ID, l.Nombre, l.Password, p.Rol FROM Usuario l JOIN Rol p ON l.ID_Rol = p.ID";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.SelectCommand = cmd;
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error botón Mostrar Usuarios:" + ex.Message);
            }
        }
        private void MostrarUsuarios_Load(object sender, EventArgs e)
        {
            mosUsuarios();
        }
    }
}
