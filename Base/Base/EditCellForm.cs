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
    public partial class EditCellForm : Form
    {
        Usuario user;
        public int valorCondicional;
        public string seleccionLocalidad;
        public string zona; 

        public EditCellForm() { }
        public EditCellForm(object valorActual, Usuario usuario, int valor)
        {
            InitializeComponent();
            textValor.Text = valorActual.ToString();
            this.user = usuario;
            this.valorCondicional=valor;
        }

        public string ObtenerNuevoValor()
        {
            // Obtén el nuevo valor ingresado por el usuario
            if (valorCondicional == 1)
            {
                //return seleccionLocalidad;
                string nuevovalor = seleccionLocalidad != null ? seleccionLocalidad.ToUpper() : null; return nuevovalor;
            }else if (valorCondicional == 2)
            {
                string nuevovalor = zona != null ? zona.ToUpper() : null; return nuevovalor;
            }
            else
            {
                //return textValor.Text;
                string nuevovalor = textValor != null ? textValor.Text.ToUpper() : null; return nuevovalor;
            }
        }

        public Usuario ObtenerNombreUsuario()
        {
            // Obtén el nombre de usuario ingresado por el usuario
            return user;
        }

        public string ObtenerComentario()
        {
            // Obtén el comentario ingresado por el usuario
            return textComentario.Text.ToUpper();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
                // Puedes agregar validación aquí si es necesario
                // Por ejemplo, asegurarte de que se hayan ingresado todos los datos requeridos.

                // Cierra el formulario con DialogResult.OK
                DialogResult = DialogResult.OK;
                Close();
        }

        private void EditCellForm_Load(object sender, EventArgs e)
        {
            if (valorCondicional == 1)
            {
                Valor2.Visible = true;
                comboBoxLocalidad.Visible = true;
                Valor1.Visible = false;
                textValor.Visible = false;
                valorZona.Visible = false;
                comboBoxZona.Visible = false;
            }else if (valorCondicional == 2)
            {
                Valor2.Visible = false;
                comboBoxLocalidad.Visible = false;
                Valor1.Visible = false;
                textValor.Visible = false;
                valorZona.Visible = true;
                comboBoxZona.Visible = true;
            }
            else
            {
                Valor2.Visible = false;
                comboBoxLocalidad.Visible=false;
                Valor1.Visible = true;
                textValor.Visible = true;
                valorZona.Visible = false;
                comboBoxZona.Visible = false;
            }

            CargarDatosEnComboBox(comboBoxLocalidad, "SELECT * FROM Lista_Localidad ORDER BY Nombre ASC", ref seleccionLocalidad);
            CargarDatosEnComboBox(comboBoxZona, "SELECT * FROM Zona", ref zona);

            comboBoxLocalidad.SelectedIndexChanged += (s, args) => seleccionLocalidad = comboBoxLocalidad.SelectedItem.ToString();
            comboBoxZona.SelectedIndexChanged += (s, args) => zona = comboBoxZona.SelectedItem.ToString();
        }

        //private void btnGuardar_Click(object sender, EventArgs e)
        //{
        //    // Puedes agregar validación aquí si es necesario
        //    // Por ejemplo, asegurarte de que se hayan ingresado todos los datos requeridos.

        //    // Cierra el formulario con DialogResult.OK
        //    DialogResult = DialogResult.OK;
        //    Close();
        //}

        //Manipulación ComboBox

        private void CargarDatosEnComboBox(System.Windows.Forms.ComboBox comboBox, string query, ref string variableSeleccionada)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();
                    SqlCommand cmd = new SqlCommand(query, conector);
                    SqlDataReader lector = cmd.ExecuteReader();

                    while (lector.Read())
                    {
                        comboBox.Items.Add(lector.GetString(1));
                    }

                    conector.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
