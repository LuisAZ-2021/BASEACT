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
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;

namespace Base
{
    public partial class submenuGeneralMostrar : Form
    {
        public submenuGeneralMostrar()
        {
            InitializeComponent();
        }

        // Declara una variable para almacenar el DataTable original para realizar Busqueda
        private DataTable dtOriginal;
        // Actualizamos las de Datos Indirectos mediante las formulas
        ActualizarFormulaIndirecto aplicarFormulas = new ActualizarFormulaIndirecto();
        public void MostrarGeneral()
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

                    string query = @"SELECT D.ID [CANTIDAD DE LOCALIDADES], D.[LOCALIDADES POR ZONA], D.ZONA, D.CÓDIGO, D.DEPARTAMENTOS, D.MUNICIPIOS, D.LOCALIDADES, D.POBLACION,
		            F.[TIPO DE FUENTE], F.[NOMBRE DE LA FUENTE], F.[NOMBRE DE ACUEDUCTO], S.[TIPO DE SERVICIO], S.[VOLUMEN DISTRIBUIDOS m3/D], S.[HS DE SERV],
		            S.[FACTURA SI/NO], S.[DOTACION ACTIVA], S.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)], S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)],
		            S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)], S.[CLOACA (SIN CONEXION - radio servido)], S.[AGUA (SIN CONEXION - radio servido)],
		            S.[AG Y CL (SIN CONEXION - radio servido)], S.[TOTAL TASA BASICA], S.[AGUA (SERVICIO MEDIDO - medidores)],
		            S.[AGUA Y CLOACA (SERVICIO MEDIDO - medidores)], S.[TOTAL SERVICIO MEDIDO], S.[TOTAL CLIENTES NO FACTURADOS], S.[TOTAL CLIENTES FACTURADOS],
		            Fi.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)], Fi.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)], Fi.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)], 
		            Fi.[CLOACA (SIN CONEXION - SERVICIO NO MEDIDO)], Fi.[AGUA (SIN CONEXION - SERVICIO NO MEDIDO)], Fi.[AG Y CL (SIN CONEXION - SERVICIO NO MEDIDO)], 
		            Fi.[TOTAL TASA BASICA], Fi.[AGUA (SERVICIO MEDIDO)], Fi.[AGUA Y CLOACA (SERVICIO MEDIDO)],
		            Fi.[TOTAL SERVICIO MEDIDO] ,Fi.[TOTAL GENERAL facturado] ,Fi.[TOTAL NO facturado],
		            I.[USUARIOS AGUA], I.[USUARIOS CLOACA], I.[USUARIOS AG Y CL], I.[TOTAL USUARIOS], I.[% COBERTURA POR CONEXION (AGUA)],
		            I.[% COBERTURA POR CONEXION (CLOACA)], I.[% COBERTURA POR CONEXION (AyC)],
		            I.[% COBERTURA POR USUARIO (AGUA)], I.[% COBERTURA POR USUARIO (CLOACA)], I.[% COBERTURA POR USUARIO (AyC)], I.[% MICROMEDICION],
		            I.[% RECAUDACION], I.[% de empleados x 1000conex]
                    FROM DatosDeLocalidades D
                    INNER JOIN dbo.DatosLocalidades_Fuente DL ON D.ID=DL.ID_DatosLocalidades
                    INNER JOIN dbo.Fuente F ON DL.ID_Fuente = F.ID
                    INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades
                    INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID
                    INNER JOIN dbo.DatosLocalidades_DatosFinancieros FL ON D.ID = FL.ID_DatosLocalidades
                    INNER JOIN dbo.DatosFinancieros Fi ON FL.ID_DatosFinancieros = Fi.ID
                    INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades
                    INNER JOIN DatosIndirectos I ON DI.ID_DatosIndirectos = I.ID;";

                    SqlCommand cmd = new SqlCommand(query, conector);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.SelectCommand = cmd;
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;

                    //Probamos la suma de los valores al final

                    // Calcular la suma de la columna "POBLACION"
                    decimal sumaLOC = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        object valor = row["POBLACION"];

                        // Verificar si el valor no es nulo y si es un número válido antes de sumarlo
                        if (valor != DBNull.Value && valor != null)
                        {
                            if (decimal.TryParse(valor.ToString(), out decimal numero))
                            {
                                sumaLOC += numero;
                            }
                        }
                    }

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
                    //totalRow["LOCALIDADES"] = "TOTAL";
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
                    totalRow["POBLACION"] = sumaLOC;

                    // Ajusta el resto de las columnas con los valores que correspondan en caso de necesitarlos

                    dt.Rows.Add(totalRow);

                    //fin de la prueba de suma

                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        column.MinimumWidth = 120; // Cambia este valor según tus necesidades (por ejemplo, 5cm)
                    }
                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al mostrar los datos de Fuente:" + ex.Message);
            }
        }
        //Probamos cheq list para seleccionar columnas

        private void submenuGeneralMostrar_Load(object sender, EventArgs e)
        {
            // Actualizamos las de Datos Indirectos mediante las formulas
            aplicarFormulas.ActualizarFormulas();
            MostrarGeneral();

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

            //Agregamos la funcionalidad para filtrar las columnas a buscar 

            // Agrega la opción "Seleccionar Todo"
            checkedListBoxColumnas.Items.Add("Seleccionar Todo");

            foreach (DataGridViewColumn columna in dataGridView1.Columns)
            {
                checkedListBoxColumnas.Items.Add(columna.Name); // Utiliza directamente el nombre de la columna
            }

            checkedListBoxColumnas.ItemCheck += checkedListBoxColumnas_ItemCheck;

            this.Controls.Add(checkedListBoxColumnas);

            // finaliza el filtrado
        }

        private void btnBusqueda_Click(object sender, EventArgs e)
        {

            //Probamos el filtro
            // Obtener la columna seleccionada y el texto de búsqueda
            string columnaBusqueda = ((OpcionCombo)comboBoxBusqueda.SelectedItem).Valor.ToString();
            string textoBusqueda = textBusqueda.Text.Trim();

            // Verificar si se seleccionó una columna y se ingresó texto de búsqueda
            if (!string.IsNullOrEmpty(columnaBusqueda) && !string.IsNullOrEmpty(textoBusqueda))
            {
                // Crear una copia del DataTable original
                DataTable dtFiltrado = dtOriginal.Copy();

                // Filtrar la copia del DataTable original basado en la columna y el texto de búsqueda
                DataView dv = new DataView(dtFiltrado);

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

                // Crear un nuevo DataTable con las columnas seleccionadas y los datos filtrados
                DataTable dtMostrar = new DataTable();
                foreach (string columna in ObtenerColumnasSeleccionadas())
                {
                    dtMostrar.Columns.Add(columna, typeof(string));
                }

                // Llenar el nuevo DataTable con los datos filtrados
                foreach (DataRowView originalRow in dv)
                {
                    DataRow newRow = dtMostrar.NewRow();
                    foreach (string columna in ObtenerColumnasSeleccionadas())
                    {
                        newRow[columna] = originalRow[columna];
                    }
                    dtMostrar.Rows.Add(newRow);
                }

                // Asignar el nuevo DataTable al DataSource del DataGridView
                dataGridView1.DataSource = dtMostrar;
            }
            else
            {
                // Si no se seleccionó una columna o no se ingresó texto de búsqueda, mostrar mensaje de advertencia
                MessageBox.Show("Seleccione una columna y ingrese texto de búsqueda.");
            }

            // Ajustar el ancho de las columnas al cargar el formulario
            AjustarAnchoColumnas();

        }

        private bool EsColumnaNumerica(DataColumn columna)
        {
            return columna.DataType == typeof(int) || columna.DataType == typeof(decimal) || columna.DataType == typeof(double);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            // Actualizamos las de Datos Indirectos mediante las formulas
            aplicarFormulas.ActualizarFormulas();
            MostrarGeneral();
            comboBoxBusqueda.SelectedIndex = 0;
            textBusqueda.Text = "";

            // Desmarcar todos los elementos en checkedListBoxColumnas
            for (int i = 0; i < checkedListBoxColumnas.Items.Count; i++)
            {
                checkedListBoxColumnas.SetItemChecked(i, false);
            }
        }

        private void checkedListBoxColumnas_ItemCheck(object sender, ItemCheckEventArgs e)
            
        {
            // Si se selecciona/deselecciona "Seleccionar Todo", aplica esa acción a todas las demás columnas
            if (e.Index == 0)
            {
                bool seleccionarTodo = (e.NewValue == CheckState.Checked);

                for (int i = 1; i < checkedListBoxColumnas.Items.Count; i++)
                {
                    checkedListBoxColumnas.SetItemChecked(i, seleccionarTodo);
                }
            }
        }

        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            // Obtener las columnas seleccionadas desde checkedListBoxColumnas
            List<string> columnasSeleccionadas = ObtenerColumnasSeleccionadas();

            // Aplicar la lógica necesaria con las columnas seleccionadas (puedes adaptar esto según tus necesidades)
            // ...

            // Después de aplicar la lógica, actualizar el DataGridView con las columnas seleccionadas
            ActualizarDataGridViewSegunColumnas(columnasSeleccionadas);
            // Ajustar el ancho de las columnas al cargar el formulario
            AjustarAnchoColumnas();
        }

        private List<string> ObtenerColumnasSeleccionadas()
        {
            List<string> columnasSeleccionadas = new List<string>();

            // Iterar sobre los elementos de checkedListBoxColumnas para obtener las columnas seleccionadas
            for (int i = 1; i < checkedListBoxColumnas.Items.Count; i++)
            {
                if (checkedListBoxColumnas.GetItemChecked(i))
                {
                    // Agregar el nombre de la columna a la lista de columnas seleccionadas
                    columnasSeleccionadas.Add(checkedListBoxColumnas.Items[i].ToString());
                }
            }

            return columnasSeleccionadas;
        }

        private void ActualizarDataGridViewSegunColumnas(List<string> columnasMostrar)
        {
            // Lógica para actualizar el DataGridView con las columnas seleccionadas

            // Crear una nueva DataTable con las columnas seleccionadas
            DataTable dtMostrar = new DataTable();
            foreach (string columna in columnasMostrar)
            {
                dtMostrar.Columns.Add(columna, typeof(string));
            }

            // Llenar la DataTable con los datos originales
            foreach (DataRow originalRow in dtOriginal.Rows)
            {
                DataRow newRow = dtMostrar.NewRow();
                foreach (string columna in columnasMostrar)
                {
                    newRow[columna] = originalRow[columna];
                }
                dtMostrar.Rows.Add(newRow);
            }

            // Asignar el nuevo DataTable al DataSource del DataGridView
            dataGridView1.DataSource = dtMostrar;
        }

        private void AjustarAnchoColumnas()
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.MinimumWidth = 120; // Cambia este valor según tus necesidades (por ejemplo, 5cm)
            }
        }

        private void btnDescargarExcel_Click(object sender, EventArgs e)
        {


            // aplico la funcionalidad sugerida

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
                    saveFile.FileName = string.Format("DatosGenerales_{0}.xlsx", DateTime.Now.ToString("ddMMyyyyHHmm"));
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
