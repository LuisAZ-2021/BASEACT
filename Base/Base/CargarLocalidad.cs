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
    public partial class CargarLocalidad : Form
    {
        public CargarLocalidad()
        {
            InitializeComponent();
        }

        private void CargarLocalidad_Load(object sender, EventArgs e)
        {
            Refresch();
        }

        public List<DatosLocalidades> listaDeLocalidades()
        {
            List<DatosLocalidades> datoslocalidades = new List<DatosLocalidades>();

            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();

                    string query = @"Select ID as [CANTIDAD LOCALIDAD], [LOCALIDADES POR ZONA], ZONA, CÓDIGO, 
                                     DEPARTAMENTOS, MUNICIPIOS, LOCALIDADES, POBLACION from dbo.DatosDeLocalidades";
                    SqlCommand cmd = new SqlCommand(query, conector);

                    //Iniciamos carga de datos de SQL
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        DatosLocalidades dat = new DatosLocalidades();
                        dat.ID = reader.IsDBNull(0) ? (int?)null : reader.GetInt32(0);
                        dat.localidadesPorZona = reader.IsDBNull(1) ? null : reader.GetString(1);
                        dat.zona = reader.IsDBNull(2) ? null : reader.GetString(2);
                        dat.codigo = reader.IsDBNull(3) ? (int?)null : reader.GetInt32(3);
                        dat.departamentos = reader.IsDBNull(4) ? null : reader.GetString(4);
                        dat.municipios = reader.IsDBNull(5) ? null : reader.GetString(5);
                        dat.localidades = reader.IsDBNull(6) ? null : reader.GetString(6);
                        dat.poblacion = reader.IsDBNull(7) ? (int?)null : reader.GetInt32(7);

                        datoslocalidades.Add(dat);
                    }

                    reader.Close();
                    conector.Close();
                    // fin Carga de datos sql


                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al extraer datos de la base:" + ex.Message);
            }

            return datoslocalidades;
        }

        public DatosLocalidades listaDeLocalidades(int Id)
        {
            DatosLocalidades dat = new DatosLocalidades();
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();

                    string query = @"Select ID as [CANTIDAD LOCALIDAD], [LOCALIDADES POR ZONA], ZONA, CÓDIGO, 
                                     DEPARTAMENTOS, MUNICIPIOS, LOCALIDADES, POBLACION from dbo.DatosDeLocalidades" +
                                     " WHERE ID=@id";
                    SqlCommand cmd = new SqlCommand(query, conector);
                    cmd.Parameters.AddWithValue("id", Id);

                    //Iniciamos carga de datos de SQL
                    SqlDataReader reader = cmd.ExecuteReader();
                    reader.Read();
                    
                    //DatosLocalidades dat = new DatosLocalidades();
                    dat.ID = reader.IsDBNull(0) ? (int?)null : reader.GetInt32(0);
                    dat.localidadesPorZona = reader.IsDBNull(1) ? null : reader.GetString(1);
                    dat.zona = reader.IsDBNull(2) ? null : reader.GetString(2);
                    dat.codigo = reader.IsDBNull(3) ? (int?)null : reader.GetInt32(3);
                    dat.departamentos = reader.IsDBNull(4) ? null : reader.GetString(4);
                    dat.municipios = reader.IsDBNull(5) ? null : reader.GetString(5);
                    dat.localidades = reader.IsDBNull(6) ? null : reader.GetString(6);
                    dat.poblacion = reader.IsDBNull(7) ? (int?)null : reader.GetInt32(7);
                    

                    reader.Close();
                    conector.Close();
                    // fin Carga de datos sql

                    return dat;
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error al extraer datos de la base:" + ex.Message);
            }

            return dat;
        }

        private void Refresch()
        {

            dataGridView1.DataSource = listaDeLocalidades();
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.MinimumWidth = 100; // Cambia este valor según tus necesidades (por ejemplo, 5cm)
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            Refresch();
        }

        private void AddLocalidades(int? id, string localidadesporZona, string zona, int? codigo, string departamentos, string municipio,
                                    string localidades, int? poblacion)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();


                    string query = "INSERT INTO dbo.DatosDeLocalidades (ID, [LOCALIDADES POR ZONA], ZONA, CÓDIGO, " +
                        "DEPARTAMENTOS, MUNICIPIOS, LOCALIDADES, POBLACION)" +
                        "VALUES (@id, @localidadesporZona, @zona, @codigo, @departamentos, @municipios, @localidades, @poblacion)";

                    SqlCommand cmd = new SqlCommand(query, conector);

                    cmd.Parameters.AddWithValue("id", id ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("localidadesporZona", string.IsNullOrEmpty(localidadesporZona) ? (object)DBNull.Value : localidadesporZona);
                    cmd.Parameters.AddWithValue("zona", string.IsNullOrEmpty(zona) ? (object)DBNull.Value : zona);
                    cmd.Parameters.AddWithValue("codigo", codigo ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("departamentos", string.IsNullOrEmpty(departamentos) ? (object)DBNull.Value : departamentos);
                    cmd.Parameters.AddWithValue("municipios", string.IsNullOrEmpty(municipio) ? (object)DBNull.Value : municipio);
                    cmd.Parameters.AddWithValue("localidades", string.IsNullOrEmpty(localidades) ? (object)DBNull.Value : localidades);
                    cmd.Parameters.AddWithValue("poblacion", poblacion ?? (object)DBNull.Value);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Localidad guardada exitosamente.");

                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocurrio un error al guardar :" + ex.Message + "El error es: "+ ex.GetType());
            }
        }


        private void EditarLocalidades(int? id, string localidadesporZona, string zona, int? codigo, string departamentos, string municipio,
                            string localidades, int? poblacion)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();


                    string query = "UPDATE dbo.DatosDeLocalidades SET [LOCALIDADES POR ZONA]=@localidadesporZona, " +
                        "ZONA=@zona, CÓDIGO=@codigo, DEPARTAMENTOS=@departamentos, MUNICIPIOS=@municipios, " +
                        "LOCALIDADES=@localidades, POBLACION=@poblacion " +
                        "WHERE ID=@id";

                    SqlCommand cmd = new SqlCommand(query, conector);

                    cmd.Parameters.AddWithValue("id", id);
                    cmd.Parameters.AddWithValue("localidadesporZona", string.IsNullOrEmpty(localidadesporZona) ? (object)DBNull.Value : localidadesporZona);
                    cmd.Parameters.AddWithValue("zona", string.IsNullOrEmpty(zona) ? (object)DBNull.Value : zona);
                    cmd.Parameters.AddWithValue("codigo", codigo ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("departamentos", string.IsNullOrEmpty(departamentos) ? (object)DBNull.Value : departamentos);
                    cmd.Parameters.AddWithValue("municipios", string.IsNullOrEmpty(municipio) ? (object)DBNull.Value : municipio);
                    cmd.Parameters.AddWithValue("localidades", string.IsNullOrEmpty(localidades) ? (object)DBNull.Value : localidades);
                    cmd.Parameters.AddWithValue("poblacion", poblacion ?? (object)DBNull.Value);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Localidad guardada exitosamente.");

                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocurrio un error al guardar :" + ex.Message);
            }
        }

        private void DeleteLocalidades(int? id)
        {
            var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
            try
            {
                using (SqlConnection conector = new SqlConnection(conString))
                {
                    conector.Open();


                    string query = "DELETE FROM dbo.DatosDeLocalidades " +
                        "WHERE ID=@id";

                    SqlCommand cmd = new SqlCommand(query, conector);

                    cmd.Parameters.AddWithValue("id", id);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Localidad borrada exitosamente.");

                    conector.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocurrio un error al eliminar :" + ex.Message);
            }
        }

        private void btnNuevo_Click(object sender, EventArgs e)
        {
            CargaLocalidadsub formcargaLocSub = new CargaLocalidadsub();

            if (formcargaLocSub.ShowDialog() == DialogResult.OK)
            {
                int? id = formcargaLocSub.ObtenerId();
                string localidadesporZona = formcargaLocSub.ObtenerLocalidadesporZona();
                string zona = formcargaLocSub.ObtenerZona();
                int? codigo = formcargaLocSub.ObtenerCodigo();
                string departamentos = formcargaLocSub.ObtenerDepartamentos();
                string municipio = formcargaLocSub.ObtenerMunicipios();
                string localidades = formcargaLocSub.ObtenerLocalidades();
                int? poblacion = formcargaLocSub.ObtenerPoblacion();

                AddLocalidades(id, localidadesporZona, zona, codigo, departamentos, municipio, localidades, poblacion);
            }
            


            Refresch();
        }

        #region Helper
        private int? GetId()
        {
            try
            {
                return int.Parse(
                    dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString()
                    );
            }
            catch
            {
                return null;
            }
        }
        #endregion

        private void btnEditar_Click(object sender, EventArgs e)
        {
            int? Id = GetId();
            if ( Id != null )
            {
                CargaLocalidadsub formEdit = new CargaLocalidadsub(Id);

                if (formEdit.ShowDialog() == DialogResult.OK)
                {
                    string localidadesporZona = formEdit.ObtenerLocalidadesporZona();
                    string zona = formEdit.ObtenerZona();
                    int? codigo = formEdit.ObtenerCodigo();
                    string departamentos = formEdit.ObtenerDepartamentos();
                    string municipio = formEdit.ObtenerMunicipios();
                    string localidades = formEdit.ObtenerLocalidades();
                    int? poblacion = formEdit.ObtenerPoblacion();

                    EditarLocalidades(Id,localidadesporZona, zona, codigo, departamentos, municipio, localidades, poblacion);
                }

                Refresch();
            }
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            int? Id = GetId();
            try
            {
                if (Id != null)
                {
                    // Mostrar un MessageBox para confirmar la eliminación
                    DialogResult resultado = MessageBox.Show("¿Estás seguro de que quieres eliminar esta fila?", "Confirmar Eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);


                    if (resultado == DialogResult.Yes)
                    {
                        CargarLocalidad objeto = new CargarLocalidad();
                        objeto.DeleteLocalidades((int)Id);
                        Refresch();
                    }
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Ocurrio un erro al eliminar",ex.Message);
            }
        }
    }


    public class DatosLocalidades
    {
        public int? ID { get; set; }
        public string localidadesPorZona{ get; set; }
        public string  zona { get; set; }

        public int? codigo { get; set; }

        public string departamentos { get; set; }

        public string municipios { get; set; }

        public string localidades { get; set; }

        public int? poblacion { get; set; }
    }
}
