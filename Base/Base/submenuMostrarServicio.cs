﻿using Base.Utilidades;
using ClosedXML.Excel;
using FontAwesome.Sharp;
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
using System.Windows.Media.Imaging;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Base
{
    public partial class submenuMostrarServicio : Form
    {
        public submenuMostrarServicio()
        {
            InitializeComponent();

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

                    string query = @"SELECT S.ID, L.LOCALIDADES, S.[TIPO DE SERVICIO], S.[VOLUMEN DISTRIBUIDOS m3/D], 
                    S.[HS DE SERV], S.[FACTURA SI/NO], S.[DOTACION ACTIVA],
		            S.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)], S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)], 
                    S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)],
		            S.[CLOACA (SIN CONEXION - radio servido)], S.[AGUA (SIN CONEXION - radio servido)], 
                    S.[AG Y CL (SIN CONEXION - radio servido)],
		            S.[TOTAL TASA BASICA], S.[AGUA (SERVICIO MEDIDO - medidores)], 
                    S.[AGUA Y CLOACA (SERVICIO MEDIDO - medidores)], S.[TOTAL SERVICIO MEDIDO],
		            S.[TOTAL CLIENTES NO FACTURADOS], S.[TOTAL CLIENTES FACTURADOS]
                    FROM DatosServicio S
                    INNER JOIN DatosLocalidades_DatosServicio DS ON DS.ID_DatosServicio = S.ID
                    INNER JOIN DatosDeLocalidades L ON DS.ID_DatosLocalidades = L.ID";
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

        private void submenuMostrarServicio_Load(object sender, EventArgs e)
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

        private string ObtenerUsuarioDelCambio(int cellID, int columna, string valorCelda)
        {
            string usuario = null;
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(conString))
            {
                connection.Open();

                // Consulta SQL para obtener el nombre de usuario del cambio
                string query = "SELECT u.Nombre FROM Cambios_DatosServicio l inner join Usuario u " +
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
                string query = "SELECT Fecha FROM Cambios_DatosServicio WHERE idCell = @CellID " +
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
                string query = "SELECT Comentario FROM Cambios_DatosServicio WHERE idCell = @CellID " +
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
                string query = "SELECT Columna FROM Cambios_DatosServicio WHERE idCell = @CellID " +
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
            MostrarServicio();
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
