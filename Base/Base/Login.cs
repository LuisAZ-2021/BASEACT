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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        public void logear(string usuarioIngresado, string passIngresada)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();


                    string query = "SELECT ID, Nombre, Password, ID_Rol FROM dbo.Usuario WHERE Nombre = @usuario AND Password = @pas";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    cmd.Parameters.AddWithValue("usuario", usuarioIngresado);
                    cmd.Parameters.AddWithValue("pas", passIngresada);

                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);

                    //Pruebo crear un usuario en forma de objeto para cargar los datos
                    Usuario usuario = new Usuario();

                    if (dt.Rows.Count == 1)
                    {
                        this.Hide();
                        object valor = 1;
                        valor = dt.Rows[0][3];
                        object id = dt.Rows[0][0];
                        usuario.ID = int.Parse(id.ToString());

                        int valorRol = int.Parse(valor.ToString());
                        usuario.Rol = valorRol;
                        usuario.Nombre = dt.Rows[0][1].ToString();



                        Inicio form = new Inicio(usuario);
                        form.Show();
                        form.FormClosing += frm_closing;

                    }
                    else
                    {
                        MessageBox.Show("Usuario y/o Contraseña Incorrecta");
                    }

                    conector.Close();

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Primer formulario:" + ex.Message);
            }
        }

        private void frm_closing(object sender, FormClosingEventArgs e)
        {
            //vaciamos los cuadros de textos
            textUsuario.Text = "";
            textPass.Text = "";
            // se muestre nuevamente el formulario de login
            this.Show();
        }
        private void btnIniciarSesion_Click(object sender, EventArgs e)
        {
            logear(this.textUsuario.Text, this.textPass.Text);
        }

        private void picMostrar_Click(object sender, EventArgs e)
        {
            //Mandamos al frente a ocultar
            picOcultar.BringToFront();
            //MOstrar contraseña
            textPass.PasswordChar = '\0';
        }

        private void picOcultar_Click(object sender, EventArgs e)
        {
            picMostrar.BringToFront();

            textPass.PasswordChar = '*';
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnMax_Click(object sender, EventArgs e)
        {
            //Si esta normal lo mazimiza, si esta maximizado vuelve normal.
            if (WindowState == FormWindowState.Normal)
            {
                WindowState = FormWindowState.Maximized;
            }
            else if (WindowState == FormWindowState.Maximized)
            {
                WindowState = FormWindowState.Normal;
            }
        }

        private void btnMin_Click(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Normal)
            {
                WindowState = FormWindowState.Minimized;
            }
            else if (WindowState == FormWindowState.Maximized)
            {
                WindowState = FormWindowState.Minimized;
            }
        }
    }
}
