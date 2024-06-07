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
    public partial class CargaLocalidadsub : Form
    {
        public int? Id;
        public CargaLocalidadsub(int? id=null)
        {
            InitializeComponent();
            this.Id = id;
            if(id != null )
                LoadData();
        }

        private void LoadData()
        {

            CargarLocalidad cargarLocalidad = new CargarLocalidad();
            DatosLocalidades datos = cargarLocalidad.listaDeLocalidades((int)Id);
            textId.Text = datos.ID.ToString();
            textLocalidadesPorZona.Text = datos.localidadesPorZona;
            textZona.Text = datos.zona;
            textCodigo.Text = datos.codigo.ToString();
            textDepartamentos.Text = datos.departamentos;
            textMunicipios.Text = datos.municipios;
            textLoc.Text = datos.localidades;
            textPoblacion.Text = datos.poblacion.ToString();
   
        }

        public int? ObtenerId()
        {
            return string.IsNullOrEmpty(textId.Text) ? (int?)null : Convert.ToInt32(textId.Text);
        }
        public string ObtenerLocalidadesporZona()
        {
            return textLocalidadesPorZona != null ? textLocalidadesPorZona.Text.ToUpper() : null;
        }

        public string ObtenerZona()
        {
            return textZona != null ? textZona.Text.ToUpper() : null;
        }

        public int? ObtenerCodigo()
        {
            return string.IsNullOrEmpty(textCodigo.Text) ? (int?)null : Convert.ToInt32(textCodigo.Text);
        }

        public string ObtenerDepartamentos()
        {
            return textDepartamentos != null ? textDepartamentos.Text.ToUpper() : null;
        }

        public string ObtenerMunicipios()
        {
            return textMunicipios != null ? textMunicipios.Text.ToUpper() : null;
        }

        public string ObtenerLocalidades()
        {
            return textLoc != null ? textLoc.Text.ToUpper() : null;
        }

        public int? ObtenerPoblacion()
        {
            return string.IsNullOrEmpty(textPoblacion.Text) ? (int?)null : Convert.ToInt32(textPoblacion.Text);
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
