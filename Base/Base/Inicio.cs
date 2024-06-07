using FontAwesome.Sharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Base
{
    public partial class Inicio : Form
    {
        public string nombre = "null";
        public int rol = 1;
        Usuario usuario;

        private static IconMenuItem MenuActivo = null;
        private static Form FormularioActivo = null;
        public Inicio(Usuario usuario)
        {
            InitializeComponent();
            lblBienvenida.Text = "¡ Hola " + usuario.Nombre + " !";
            if(usuario.Rol == 1)
            {
                lblRol.Text = "Superusuario";
            }else if(usuario.Rol == 2)
            {
                lblRol.Text = "Administrador";
            } else if(usuario.Rol == 3)
            {
                lblRol.Text = "Usuario";
            }
            this.rol = usuario.Rol;
            this.usuario = usuario;
        }

        private void AbrirFormulario(IconMenuItem menu, Form formulario)
        {
            if (MenuActivo != null)
            {
                MenuActivo.BackColor = Color.SkyBlue;
            }

            menu.BackColor = Color.Silver;
            MenuActivo = menu;

            if (FormularioActivo != null)
            {
                FormularioActivo.Close();
            }

            FormularioActivo = formulario;
            formulario.TopLevel = false;
            formulario.FormBorderStyle = FormBorderStyle.None;
            formulario.Dock = DockStyle.Fill;
            formulario.BackColor = Color.LightCyan;

            contenedor.Controls.Add(formulario);
            formulario.Show();

        }

        private void Inicio_Load(object sender, EventArgs e)
        {
            if (rol == 3)  
            {
                submenuCargarLocalidad.Visible = false;
                submenuActualizarLocalidad.Visible = false; 
                menuUsuarios.Visible = false;
                submenuActualizarFuente.Visible = false;
                submenuActualizarServicio.Visible = false;
                submenuActualizarFinanciero.Visible = false;
                submenuActualizarIndirecto.Visible = false;
            }
            else if (rol == 2)
            {
                menuUsuarios.Visible = false;
            }
        }

        private void submenuMostrarLocalidad_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuLocalidad, new DetalleLocalidad());
        }

        private void submenuCargarLocalidad_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuLocalidad, new CargarLocalidad());
        }

        private void submenuActualizarLocalidad_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuLocalidad, new ActualizarLocalidad(usuario));
        }

        private void submenuMostrarUsuarios_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuUsuarios, new MostrarUsuarios());
        }

        private void submenuCargarUsuarios_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuUsuarios, new CargarUsuarios());
        }

        private void submenuMostrarFuente_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuFuente, new subMenuMostrarFuente());
        }

        private void submenuMostrarGeneral_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuGeneral, new submenuGeneralMostrar());
        }

        private void submenuActualizarFuente_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuFuente, new ActualizarFuente(usuario));
        }

        private void submenuMostrarServicio_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuServicio, new submenuMostrarServicio());
        }

        private void submenuActualizarServicio_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuServicio, new ActualizarServicio(usuario));
        }

        private void submenuMostrarFinanciero_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuFinanciero, new submenuMostrarFinanciero());
        }

        private void submenuActualizarFinanciero_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuFinanciero, new ActualizarFinanciero(usuario));
        }

        private void submenuMostrarIndirecto_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuIndirecto, new submenuMostrarIndirecto());
        }

        private void submenuActualizarIndirecto_Click(object sender, EventArgs e)
        {
            AbrirFormulario(menuIndirecto, new ActualizarIndirecto(usuario));
        }
    }
}
