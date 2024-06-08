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
    public partial class subMenuMostrarFuente : Form
    {
        public subMenuMostrarFuente()
        {
            InitializeComponent();
        }

        // Declara una variable para almacenar el DataTable original para realizar Busqueda
        private DataTable dtOriginal;
        public void MostrarFuente()
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

                    string query = @"SELECT F.ID, D.LOCALIDADES, F.[TIPO DE FUENTE], F.[NOMBRE DE LA FUENTE],F.[NOMBRE DE ACUEDUCTO]
                                    FROM DatosDeLocalidades D
                                    INNER JOIN DatosFuente F ON D.ID = F.ID_LOCALIDAD";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.SelectCommand = cmd;
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

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

                MessageBox.Show("Error al mostrar los datos de Fuente:" + ex.Message);
            }
        }

        private void subMenuMostrarFuente_Load(object sender, EventArgs e)
        {
            MostrarFuente();

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

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            MostrarFuente();
            comboBoxBusqueda.SelectedIndex = 0;
            textBusqueda.Text = "";
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
                    saveFile.FileName = string.Format("DatosFuente_{0}.xlsx", DateTime.Now.ToString("ddMMyyyyHHmm"));
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
