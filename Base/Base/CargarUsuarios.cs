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
    public partial class CargarUsuarios : Form

    {
        // Agregar una variable global para almacenar el ID del usuario seleccionado
        private int usuarioIdSeleccionado;
        private bool edicionHabilitada = false;
        public CargarUsuarios()
        {
            InitializeComponent();

            btnActualizar.Enabled = false;
        }

        //Función para cargar los roles
        private List<RolItem> roles = new List<RolItem>();
        int rolId;
        private void CargarRoles()
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();
                    string query = "SELECT ID, Rol FROM Rol";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    SqlDataReader reader = cmd.ExecuteReader();

                    // Antes de agregar los roles, limpiar el ComboBox
                    comboBoxRol.Items.Clear();

                    while (reader.Read())
                    {
                        int rolId = (int)reader["ID"];
                        string rolNombre = reader["Rol"].ToString();
                        // Crea un objeto RolItem y agrégalo a la lista.
                        RolItem rol = new RolItem(rolId, rolNombre);
                        roles.Add(rol);
                        comboBoxRol.Items.Add(rol.Rol);
                    }

                    reader.Close();
                    conector.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar los roles: " + ex.Message);
            }
        }
        //Selección del rol
        private void comboBoxRoles_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxRol.SelectedIndex >= 0)
            {
                RolItem rolSeleccionado = roles[comboBoxRol.SelectedIndex];
                rolId = rolSeleccionado.Id;
                // Puedes usar rolId como la clave foránea cuando creas el usuario.
            }
        }


        public void crearUsuario(string nombre, string pass, int rol, bool esEdicion = false)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();

                    /*
                    INSERT INTO Usuario(Nombre, Password, Rol)
                    VALUES(@nombre, @pass, @Rol)*/

                    /*string query = "INSERT INTO Usuario(Nombre, Password, Rol)" +
                        "VALUES(@nombre, @password, @rol)";*/
                    string query;
                    if (esEdicion)
                    {
                        // Si es una edición, actualizar el usuario existente
                        query = "UPDATE Usuario SET Nombre=@nombre, Password=@password, ID_Rol=@rol WHERE ID=@userId";
                    }
                    else
                    {
                        // Si es una creación, insertar un nuevo usuario
                        query = "INSERT INTO Usuario(Nombre, Password, ID_Rol) VALUES(@nombre, @password, @rol)";
                    }

                    SqlCommand cmd = new SqlCommand(query, conector);
                    cmd.Parameters.AddWithValue("nombre", nombre);
                    cmd.Parameters.AddWithValue("password", pass);
                    cmd.Parameters.AddWithValue("rol", rol);

                    // Si es una edición, agregar el parámetro del ID del usuario
                    if (esEdicion)
                    {
                        cmd.Parameters.AddWithValue("userId", usuarioIdSeleccionado);
                    }


                    cmd.ExecuteNonQuery();

                    //MessageBox.Show("Usuario guardado exitosamente.");
                    MessageBox.Show(esEdicion ? "Usuario actualizado exitosamente." : "Usuario creado exitosamente.");

                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al guardar un Usuario:" + ex.Message);
            }
        }

        private void btnCrearUsuario_Click(object sender, EventArgs e)
        {
            //string nombre = "";
            //string password = "";
            //string clave = "";
            //crearUsuario(nombre, password, clave);
            string nombre = textBoxNombre.Text; // Obtener el nombre desde un TextBox.
            string password = textBoxPass.Text; // Obtener la contraseña desde un TextBox.

            // Obtener el rol seleccionado del ComboBox.
            if (rolId != 0)
            {

                // Llamar a la función para crear un usuario con el rol seleccionado.
                crearUsuario(nombre, password, rolId);
                mosUsuarios();
                // Limpiar los campos y recargar la lista de usuarios
                textBoxNombre.Text = "";
                textBoxPass.Text = "";
                comboBoxRol.SelectedIndex = -1;
                usuarioIdSeleccionado = 0;

            }
            else
            {
                MessageBox.Show("Por favor, seleccione un rol.");
            }
        }

        private void iconGenerarClave_Click(object sender, EventArgs e)
        {
            string contrasenaGenerada = GenerarContrasenaAleatoria();
            textBoxPass.Text = contrasenaGenerada;
        }

        private string GenerarContrasenaAleatoria()
        {
            string caracteres = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            Random random = new Random();
            int longitud = 6; // Cambia la longitud de la contraseña según tus preferencias.
            char[] contrasena = new char[longitud];

            for (int i = 0; i < longitud; i++)
            {
                contrasena[i] = caracteres[random.Next(caracteres.Length)];
            }

            return new string(contrasena);
        }

        private void CargarUsuarios_Load(object sender, EventArgs e)
        {
            CargarRoles(); // Llama a la función para cargar los roles en el ComboBox.
            comboBoxRol.SelectedIndexChanged += comboBoxRoles_SelectedIndexChanged;
            // Establecer el color inicial del botón
            btnEditar.BackColor = SystemColors.Control;

            // muestro en el data grid los usuarios
            mosUsuarios();
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

                    //Pruebo datagridView para conectar a los usuarios
                    dataGridUsuarios.DataSource = dt;
                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error botón Mostrar Usuarios:" + ex.Message);
            }
        }

        private void DeleteUsuario(int? id)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();


                    string query = "DELETE FROM dbo.Usuario " +
                        "WHERE ID=@id";

                    SqlCommand cmd = new SqlCommand(query, conector);

                    cmd.Parameters.AddWithValue("id", id);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Usuario borrado exitosamente.");

                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocurrio un error al eliminar :" + ex.Message);
            }
        }

        #region Helper
        private int? GetId()
        {
            try
            {
                return int.Parse(
                    dataGridUsuarios.Rows[dataGridUsuarios.CurrentRow.Index].Cells[0].Value.ToString()
                    );
            }
            catch
            {
                return null;
            }
        }
        #endregion

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            int? Id = GetId();
            try
            {
                if (Id != null)
                {
                    string name = dataGridUsuarios.Rows[dataGridUsuarios.CurrentRow.Index].Cells[1].Value.ToString();
                    // Mostrar un MessageBox para confirmar la eliminación
                    DialogResult resultado = MessageBox.Show("¿Estás seguro de que quieres eliminar a "+name+"?", "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);


                    if (resultado == DialogResult.Yes)
                    {
                        //CargarLocalidad objeto = new CargarLocalidad();
                        //objeto.DeleteLocalidades((int)Id);
                        DeleteUsuario((int)Id);
                        mosUsuarios();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un error al eliminar", ex.Message);
            }
        }

        
        private void btnEditar_Click(object sender, EventArgs e)
        {
            // Cambiar el estado de la edición
            edicionHabilitada = !edicionHabilitada;

            // Habilitar o deshabilitar el evento CellClick según el estado
            if (edicionHabilitada)
            {
                //dataGridUsuarios.CellClick += dataGridUsuarios_CellClick;
                dataGridUsuarios.CellClick += ediciondeUsuarios;
            }
            else
            {
                //dataGridUsuarios.CellClick -= dataGridUsuarios_CellClick;
                dataGridUsuarios.CellClick -= ediciondeUsuarios;
            }

            // Cambiar el color del botón según el estado
            btnEditar.BackColor = edicionHabilitada ? Color.Green : SystemColors.Control;
            btnEliminar.Enabled = !edicionHabilitada;
            btnCrearUsuario.Enabled = !edicionHabilitada;
            // Cambiar la habilitación del botón "Actualizar"
            btnActualizar.Enabled = edicionHabilitada;

            // Otros controles relacionados con la edición...

            // Actualizar el texto del botón
            btnEditar.Text = edicionHabilitada ? "Finalizar Edición" : "EDITAR";

            //probamos limpiar los campos

            if(edicionHabilitada == false)
            {
                // Limpiar los campos y recargar la lista de usuarios
                textBoxNombre.Text = "";
                textBoxPass.Text = "";
                comboBoxRol.SelectedIndex = -1;
                usuarioIdSeleccionado = 0;
            }
        
        }

        private void ediciondeUsuarios(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Verificar si se ha seleccionado un usuario para editar
                if (e.RowIndex >= 0)
                {
                    DataGridViewRow row = dataGridUsuarios.Rows[e.RowIndex];

                    // Obtener la información del usuario seleccionado
                    int userId = Convert.ToInt32(row.Cells["Id"].Value);
                    string userName = row.Cells["Nombre"].Value.ToString();
                    string userPassword = row.Cells["Password"].Value.ToString();
                    string userRol = row.Cells["Rol"].Value.ToString();

                    // Mostrar la información del usuario en los campos correspondientes
                    textBoxNombre.Text = userName;
                    textBoxPass.Text = userPassword;

                    // Seleccionar el rol del usuario en el ComboBox
                    comboBoxRol.SelectedIndex = comboBoxRol.FindStringExact(userRol);

                    // Guardar el ID del usuario seleccionado en una variable global para su posterior uso
                    usuarioIdSeleccionado = userId;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No seleccionaste correctamente, vuelve a intentarlo. ");
            }
        }

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            if (usuarioIdSeleccionado != 0)
            {
                // Obtener la información actualizada del formulario
                string nombre = textBoxNombre.Text;
                string password = textBoxPass.Text;

                // Obtener el rol seleccionado del ComboBox
                int rolId = roles[comboBoxRol.SelectedIndex].Id;

                // Llamar a la función para editar el usuario existente
                crearUsuario(nombre, password, rolId, esEdicion: true);

                // Limpiar los campos y recargar la lista de usuarios
                textBoxNombre.Text = "";
                textBoxPass.Text = "";
                comboBoxRol.SelectedIndex = -1;
                usuarioIdSeleccionado = 0;
                mosUsuarios();
            }
            else
            {
                MessageBox.Show("Por favor, selecciona un usuario para editar.");
            }
        }
    }
}
