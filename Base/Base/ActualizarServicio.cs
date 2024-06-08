using Base.Utilidades;
using ClosedXML.Excel;
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
    public partial class ActualizarServicio : Form
    {
        Usuario user;
        public ActualizarServicio(Usuario user)
        {
            InitializeComponent();
            this.user = user;
        }

        // Declara una variable para almacenar el DataTable original para realizar Busqueda
        private DataTable dtOriginal;

        public void MostrarServicio()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();
                    DataTable dt = new DataTable();

                    string query = @"select S.ID, L.LOCALIDADES, S.[TIPO DE SERVICIO], S.[VOLUMEN DISTRIBUIDOS m3/D], S.[HS DE SERV],
		                            S.[FACTURA SI/NO], S.[DOTACION ACTIVA], S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)], S.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)], 
		                            S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)], S.[AGUA (SIN CONEXION - radio servido)], S.[CLOACA (SIN CONEXION - radio servido)],
		                            S.[AG Y CL (SIN CONEXION - radio servido)], S.[TOTAL TASA BASICA], S.[AGUA (SERVICIO MEDIDO - medidores)],
		                            S.[AGUA Y CLOACA (SERVICIO MEDIDO - medidores)], S.[TOTAL SERVICIO MEDIDO],
		                            S.[TOTAL CLIENTES NO FACTURADOS], S.[TOTAL CLIENTES FACTURADOS]
                                    from DatosServicio S
                                    inner join DatosDeLocalidades L ON S.ID_LOCALIDAD=L.ID";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.SelectCommand = cmd;
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    //Probamos la suma de los valores al final

                    // Calcular la suma de la columna "VOLUMEN DISTRIBUIDOS m3/D"
                    decimal sumaVD = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["VOLUMEN DISTRIBUIDOS m3/D"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaVD += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "HS DE SERV"
                    decimal sumaHS = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["HS DE SERV"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaHS += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "DOTACION ACTIVA"
                    decimal sumaDA = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["DOTACION ACTIVA"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaDA += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "AGUA (CON CONEXION - SERVICIO NO MEDIDO)"
                    decimal sumaACCSNM = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["AGUA (CON CONEXION - SERVICIO NO MEDIDO)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaACCSNM += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "CLOACA (CON CONEXION - SERVICIO NO MEDIDO)"
                    decimal sumaCCSNM = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["CLOACA (CON CONEXION - SERVICIO NO MEDIDO)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaCCSNM += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)"
                    decimal sumaAyCCCSNM = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaAyCCCSNM += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "AGUA (SIN CONEXION - radio servido)"
                    decimal sumaASC = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["AGUA (SIN CONEXION - radio servido)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaASC += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "CLOACA (SIN CONEXION - radio servido)"
                    decimal sumaCSC = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["CLOACA (SIN CONEXION - radio servido)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaCSC += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "AG Y CL (SIN CONEXION - radio servido)"
                    decimal sumaAyC = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["AG Y CL (SIN CONEXION - radio servido)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaAyC += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "TOTAL TASA BASICA"
                    decimal sumaTTB = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["TOTAL TASA BASICA"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaTTB += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "AGUA (SERVICIO MEDIDO - medidores)"
                    decimal sumaASM = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["AGUA (SERVICIO MEDIDO - medidores)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaASM += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "AGUA Y CLOACA (SERVICIO MEDIDO - medidores)"
                    decimal sumaAyCSM = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["AGUA Y CLOACA (SERVICIO MEDIDO - medidores)"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaAyCSM += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "TOTAL SERVICIO MEDIDO"
                    decimal sumaTSM = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["TOTAL SERVICIO MEDIDO"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaTSM += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "TOTAL CLIENTES NO FACTURADOS"
                    decimal sumaTCNF = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["TOTAL CLIENTES NO FACTURADOS"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaTCNF += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Calcular la suma de la columna "TOTAL CLIENTES FACTURADOS"
                    decimal sumaTCF = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["TOTAL CLIENTES FACTURADOS"];
                        if (valor != DBNull.Value && valor != null)
                        {
                            sumaTCF += Convert.ToDecimal(valor);
                        }
                        // Si prefieres asignar un valor predeterminado a los null, puedes hacer algo así:
                        // suma += valor != DBNull.Value && valor != null ? Convert.ToDecimal(valor) : 0;
                    }

                    // Agregar una fila al DataTable con la suma calculada
                    DataRow totalRow = dt.NewRow();
                    totalRow["LOCALIDADES"] = "TOTAL";
                    totalRow["VOLUMEN DISTRIBUIDOS m3/D"] = sumaVD;
                    totalRow["HS DE SERV"] = sumaHS;
                    totalRow["DOTACION ACTIVA"] = sumaDA;
                    totalRow["AGUA (CON CONEXION - SERVICIO NO MEDIDO)"] = sumaACCSNM;
                    totalRow["CLOACA (CON CONEXION - SERVICIO NO MEDIDO)"] = sumaCCSNM;
                    totalRow["AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)"] = sumaAyCCCSNM;
                    totalRow["AGUA (SIN CONEXION - radio servido)"] = sumaASC;
                    totalRow["CLOACA (SIN CONEXION - radio servido)"] = sumaCSC;
                    totalRow["AG Y CL (SIN CONEXION - radio servido)"] = sumaAyC;
                    totalRow["TOTAL TASA BASICA"] = sumaTTB;
                    totalRow["AGUA (SERVICIO MEDIDO - medidores)"] = sumaASM;
                    totalRow["AGUA Y CLOACA (SERVICIO MEDIDO - medidores)"] = sumaAyCSM;
                    totalRow["TOTAL SERVICIO MEDIDO"] = sumaTSM;
                    totalRow["TOTAL CLIENTES NO FACTURADOS"] = sumaTCNF;
                    totalRow["TOTAL CLIENTES FACTURADOS"] = sumaTCF;

                    // Ajusta el resto de las columnas con los valores que correspondan en caso de necesitarlos

                    dt.Rows.Add(totalRow);


                    //fin de la prueba de suma

                    // Oculta la columna de ID en el DataGridView

                    dataGridView1.Columns["ID"].Visible = false; // Reemplaza "ID" 

                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        column.MinimumWidth = 100; // Cambia este valor según tus necesidades (por ejemplo, 5cm)
                    }
                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al mostrar los datos de Servicio:" + ex.Message);
            }
        }

        private void ActualizarServicio_Load(object sender, EventArgs e)
        {
            MostrarServicio();

            // Obtén el DataTable original y asígnalo a la variable dtOriginal para la Busqueda
            dtOriginal = (DataTable)dataGridView1.DataSource;

            foreach (DataGridViewColumn columna in dataGridView1.Columns)
            {
                if (columna.Visible == true)
                {
                    comboBoxBusqueda.Items.Add(new OpcionCombo() { Valor = columna.Name, Texto = columna.HeaderText });
                }
            }

            comboBoxBusqueda.DisplayMember = "Texto";
            comboBoxBusqueda.ValueMember = "Valor";
            comboBoxBusqueda.SelectedIndex = 0;
        }

        private void UpdateLocTable(int cellID, string newValue)
        {
            try
            {

                if (dataGridView1.CurrentCell != null)
                {
                    string columnName = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].Name;
                    var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
                    using (SqlConnection connection = new SqlConnection(conString))
                    {
                        connection.Open();
                        string updateQuery = $"UPDATE DatosServicio SET [{columnName}] = @NewValue WHERE ID = @CellID";
                        //Agregamos la tabla para la suma de las columnas anteriores
                        string updateQueryTTB = $"UPDATE DatosServicio SET [TOTAL TASA BASICA] = " +
                            $"ISNULL([AGUA (CON CONEXION - SERVICIO NO MEDIDO)],0) + ISNULL([CLOACA (CON CONEXION - SERVICIO NO MEDIDO)],0) +" +
                            $" ISNULL([AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)],0) + ISNULL([AGUA (SIN CONEXION - radio servido)],0) + " +
                            $"ISNULL([CLOACA (SIN CONEXION - radio servido)],0) + ISNULL([AG Y CL (SIN CONEXION - radio servido)],0)";

                        string updateQuerySM = $"UPDATE DatosServicio SET [TOTAL SERVICIO MEDIDO] = " +
                            $"ISNULL([AGUA (SERVICIO MEDIDO - medidores)],0) + ISNULL([AGUA Y CLOACA (SERVICIO MEDIDO - medidores)],0)";

                        string updateQueryCF = $"UPDATE DatosServicio SET [TOTAL CLIENTES FACTURADOS] = " +
                            $"ISNULL([TOTAL TASA BASICA],0) + ISNULL([TOTAL SERVICIO MEDIDO],0) - ISNULL([TOTAL CLIENTES NO FACTURADOS],0)";

                        SqlCommand commandTTB = new SqlCommand(updateQueryTTB, connection);
                        SqlCommand commandSM = new SqlCommand(updateQuerySM, connection);
                        SqlCommand commandCF = new SqlCommand(updateQueryCF, connection);
                        //fin de la suma
                        SqlCommand command = new SqlCommand(updateQuery, connection);

                        if (string.IsNullOrEmpty(newValue)) // Si el valor es en blanco
                        {
                            command.Parameters.AddWithValue("@NewValue", DBNull.Value);
                        }
                        else
                        {
                            command.Parameters.AddWithValue("@NewValue", newValue);
                        }
                        command.Parameters.AddWithValue("@CellID", cellID);
                        command.ExecuteNonQuery();
                        //Inicio suma
                        commandTTB.ExecuteNonQuery();
                        commandSM.ExecuteNonQuery();
                        commandCF.ExecuteNonQuery();
                        //FIN suma
                    }
                }
                else
                {
                    // Manejo de error o mensaje al usuario si no hay una celda actualmente seleccionada.
                    MessageBox.Show("No hay celda seleccionada");
                }
            }
            catch (SqlException ex)
            {
                // Manejar la excepción específica de SQL
                MessageBox.Show($"Error al actualizar la base de datos: {ex.Message}");
                // Puedes registrar o manejar la excepción de alguna otra manera aquí
            }
            catch (Exception ex)
            {
                // Manejar otras excepciones no específicas de SQL
                MessageBox.Show($"Error inesperado: {ex.Message}");
                // Puedes registrar o manejar la excepción de alguna otra manera aquí
            }
        }
        int idTablaModificada = 3;
        private void InsertChangeRecord(int cellID, int columna, int userID, DateTime changeDate, string comment, string valorCelda, int tabla)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();
                string insertQuery = "INSERT INTO Cambios (idCell, Columna, ValorCelda, Usuario, Fecha, Comentario, Id_Tabla) " +
                    "VALUES (@CellID, @Columna, @ValorCelda, @Userid, @ChangeDate, @Comment, @tabla)";
                SqlCommand command = new SqlCommand(insertQuery, connection);
                command.Parameters.AddWithValue("@CellID", cellID);
                command.Parameters.AddWithValue("@Columna", columna);
                //command.Parameters.AddWithValue("@ValorCelda", valorCelda ?? (object)DBNull.Value);
                if (string.IsNullOrEmpty(valorCelda)) // Si el valor es en blanco
                {
                    command.Parameters.AddWithValue("@ValorCelda", DBNull.Value);
                }
                else
                {
                    command.Parameters.AddWithValue("@ValorCelda", valorCelda);
                }
                command.Parameters.AddWithValue("@Userid", userID);
                command.Parameters.AddWithValue("@ChangeDate", changeDate);
                command.Parameters.AddWithValue("@Comment", comment);
                command.Parameters.AddWithValue("@tabla", tabla);
                command.ExecuteNonQuery();
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                // Obtén el valor actual de la celda
                //int cellID = (int)dataGridView1.Rows[e.RowIndex].Cells["ID"].Value;
                int cellID = (int)dataGridView1.Rows[e.RowIndex].Cells["ID"].Value;

                string valorActual = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                // Crea una instancia del formulario personalizado
                //int condicional = 1;
                EditCellForm editForm;
                //Prueba de desplegable 
                if (dataGridView1.Columns[e.ColumnIndex].Name == "LOCALIDADES")
                {
                    // No asignar un formulario, simplemente mostrar el mensaje
                    editForm = null;
                    MessageBox.Show("No puedes modificar la localidad, para ello tienes que ir a la pestaña de Localidad");
                }
                else if (dataGridView1.Columns[e.ColumnIndex].Name == "ZONA")
                {
                    editForm = new EditCellForm(valorActual, user, 2);
                }
                else
                {
                    editForm = new EditCellForm(valorActual, user, 0);
                }

                //Fin de desplegable



                // Muestra el formulario al usuario
                if (editForm != null && editForm.ShowDialog() == DialogResult.OK)
                {
                    // Cuando el usuario hace clic en "Aceptar" en el formulario
                    // Obtén el nuevo valor, nombre de usuario y comentario
                    string nuevoValor = editForm.ObtenerNuevoValor();
                    Usuario user = editForm.ObtenerNombreUsuario();
                    string comment = editForm.ObtenerComentario();

                    // Actualiza la base de datos y la cuadrícula con los nuevos valores
                    // ...
                    UpdateLocTable(cellID, nuevoValor);
                    DateTime changeDate = DateTime.Now;
                    //purebo acceder al indice de la columna

                    int columna = e.ColumnIndex;

                    // actualizo para ingresar el valor de celda
                    InsertChangeRecord(cellID, columna, user.ID, changeDate, comment, nuevoValor, idTablaModificada);

                    // Actualiza la cuadrícula con el nuevo valor
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = nuevoValor;



                    if (string.IsNullOrEmpty(nuevoValor))
                    {
                        // Si nuevoValor es nulo o una cadena vacía, asigna DBNull.Value a la celda
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = DBNull.Value;
                    }
                    else
                    {
                        // Si nuevoValor no está vacío, actualiza la celda con el nuevo valor
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = nuevoValor;
                    }
                }
            }
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                //int cellID = (int)dataGridView1.Rows[e.RowIndex].Cells["ID"].Value;
                int cellID = (int)dataGridView1.Rows[e.RowIndex].Cells["ID"].Value;

                // Obtener el valor de la celda seleccionada
                string valorCelda = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

                // Obtener la información de la columna modificada
                int columna = ObtenerColumnaModificada(cellID, valorCelda);

                // Resto del código para mostrar información
                string usuario = ObtenerUsuarioDelCambio(cellID, columna, valorCelda);
                DateTime fecha = ObtenerFechaDelCambio(cellID, columna, valorCelda);
                string comentario = ObtenerComentarioDelCambio(cellID, columna, valorCelda);

                MessageBox.Show($"Usuario: {usuario}\nFecha: {fecha}\nFuente: {comentario}");

                //#####################################################################

            }
        }

        private string ObtenerUsuarioDelCambio(int cellID, int columna, string valorCelda)
        {
            string usuario = null;
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();

                // Consulta SQL para obtener el nombre de usuario del cambio
                string query = "SELECT u.Nombre FROM Cambios l inner join Usuario u " +
                    "on u.Id = l.Usuario " +
                    "WHERE idCell = @CellID " +
                    "AND Columna = @Columna AND ValorCelda = @ValorCelda";



                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@CellID", cellID);
                    command.Parameters.AddWithValue("@Columna", columna);
                    command.Parameters.AddWithValue("@ValorCelda", valorCelda);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            // Obtenemos el nombre de usuario de la columna "Usuario"
                            usuario = reader["Nombre"].ToString();
                        }
                    }
                }
            }

            return usuario;
        }

        private DateTime ObtenerFechaDelCambio(int cellID, int columna, string valorCelda)
        {
            DateTime fecha = DateTime.MinValue; // Valor predeterminado en caso de no encontrar la fecha.

            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();

                // Consulta SQL para obtener la fecha del cambio
                string query = "SELECT Fecha FROM Cambios WHERE idCell = @CellID " +
                    "AND Columna = @Columna AND ValorCelda = @ValorCelda";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@CellID", cellID);
                    command.Parameters.AddWithValue("@Columna", columna);
                    command.Parameters.AddWithValue("@ValorCelda", valorCelda);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            // Obtenemos la fecha de la columna "Fecha"
                            fecha = Convert.ToDateTime(reader["Fecha"]);
                        }
                    }
                }
            }

            return fecha;
        }

        private string ObtenerComentarioDelCambio(int cellID, int columna, string valorCelda)
        {
            string comentario = string.Empty; // Valor predeterminado en caso de no encontrar un comentario.

            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();

                // Consulta SQL para obtener el comentario del cambio
                string query = "SELECT Comentario FROM Cambios WHERE idCell = @CellID " +
                    "AND Columna = @Columna AND ValorCelda = @ValorCelda";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@CellID", cellID);
                    command.Parameters.AddWithValue("@Columna", columna);
                    command.Parameters.AddWithValue("@ValorCelda", valorCelda);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            // Obtenemos el comentario de la columna "Comentario"
                            comentario = reader["Comentario"].ToString();
                        }
                    }
                }
            }

            return comentario;
        }

        private int ObtenerColumnaModificada(int cellID, string valorCelda)
        {
            int columna = -1; // Valor predeterminado en caso de no encontrar la columna.

            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();

                // Consulta SQL para obtener la columna modificada
                string query = "SELECT Columna FROM Cambios WHERE idCell = @CellID " +
                    "AND ValorCelda = @ValorCelda";

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@CellID", cellID);
                    command.Parameters.AddWithValue("@ValorCelda", valorCelda);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            // Obtenemos el índice de la columna de la columna "Columna"
                            columna = Convert.ToInt32(reader["Columna"]);
                        }
                    }
                }
            }

            return columna;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            MostrarServicio();
            comboBoxBusqueda.SelectedIndex = 0;
            textBusqueda.Text = "";
        }

        private void btnSincronizarDatos_Click(object sender, EventArgs e)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();

                    string query = @"EXEC InsertarDatosIterativo_DatosServicios;";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    cmd.ExecuteNonQuery();
                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al Sincronizar las Localidades:" + ex.Message);
            }

            MostrarServicio();
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {
            // Obtener la columna seleccionada y el texto de búsqueda
            string columnaBusqueda = ((OpcionCombo)comboBoxBusqueda.SelectedItem).Valor.ToString();
            string textoBusqueda = textBusqueda.Text.Trim();

            // Verificar si se seleccionó una columna y se ingresó texto de búsqueda
            if (!string.IsNullOrEmpty(columnaBusqueda) && !string.IsNullOrEmpty(textoBusqueda))
            {
                // Obtener el DataTable actual del DataGridView
                DataTable dtActual = (DataTable)dataGridView1.DataSource;


                // Filtrar el DataTable basado en la columna y el texto de búsqueda

                DataView dv = new DataView(dtOriginal);

                // Construir la cláusula RowFilter envolviendo el nombre de la columna con corchetes
                string columnaFiltrado = $"[{columnaBusqueda}]";

                // Intentar convertir el valor de búsqueda a un número
                if (EsColumnaNumerica(dtOriginal.Columns[columnaBusqueda]))
                {
                    if (double.TryParse(textoBusqueda, out double valorNumerico))
                    {
                        // Filtrar números usando Equals
                        dv.RowFilter = $"{columnaFiltrado} = {valorNumerico}";
                    }
                    else
                    {
                        try
                        {
                            // Si la conversión falla, buscar por igualdad de texto
                            dv.RowFilter = $"{columnaFiltrado} = '{textoBusqueda}'";
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Tienes que ingresar un número: ");
                        }
                    }
                }
                else
                {
                    // Filtrar texto usando LIKE
                    dv.RowFilter = $"{columnaFiltrado} LIKE '%{textoBusqueda}%'";
                }


                // Asignar el nuevo DataTable filtrado al DataSource del DataGridView
                dataGridView1.DataSource = dv.ToTable();
            }
            else
            {
                // Si no se seleccionó una columna o no se ingresó texto de búsqueda, mostrar mensaje de advertencia
                MessageBox.Show("Seleccione una columna y ingrese texto de búsqueda.");
            }
        }

        private bool EsColumnaNumerica(DataColumn columna)
        {
            return columna.DataType == typeof(int) || columna.DataType == typeof(decimal) || columna.DataType == typeof(double);
        }

        private void btnDescargarExcel_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count < 1)
            {
                MessageBox.Show("No hay datos para Exportar", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                try
                {
                    // Crear un DataTable para almacenar los datos a exportar
                    DataTable dtExportar = new DataTable();

                    // Agregar las columnas al DataTable
                    foreach (DataGridViewColumn columna in dataGridView1.Columns)
                    {
                        if (columna.HeaderText != "")
                        {
                            dtExportar.Columns.Add(columna.HeaderText, typeof(string));
                        }
                    }

                    // Agregar las filas visibles al DataTable
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Visible)
                        {
                            DataRow newRow = dtExportar.NewRow();

                            foreach (DataGridViewColumn columna in dataGridView1.Columns)
                            {
                                if (columna.HeaderText != "")
                                {
                                    newRow[columna.HeaderText] = row.Cells[columna.Index].Value;
                                }
                            }

                            dtExportar.Rows.Add(newRow);
                        }
                    }

                    // Mostrar el cuadro de diálogo para guardar el archivo Excel
                    SaveFileDialog saveFile = new SaveFileDialog();
                    saveFile.FileName = string.Format("DatosServicio_{0}.xlsx", DateTime.Now.ToString("ddMMyyyyHHmm"));
                    saveFile.Filter = "Excel File | *.xlsx";

                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {
                        // Guardar el DataTable en un archivo Excel utilizando la librería ClosedXML
                        using (var workbook = new XLWorkbook())
                        {
                            workbook.Worksheets.Add(dtExportar, "Informe");
                            workbook.SaveAs(saveFile.FileName);
                        }

                        MessageBox.Show("Datos descargados exitosamente", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al descargar los datos: {ex.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
