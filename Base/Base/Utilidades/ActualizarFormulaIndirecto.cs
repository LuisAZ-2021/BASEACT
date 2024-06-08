using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.VariantTypes;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;

namespace Base.Utilidades
{
    public class ActualizarFormulaIndirecto
    {
        public void ActualizarFormulas()
        {
            try
            {
                var conString = ConfigurationManager.ConnectionStrings["dbSql"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(conString))
                {
                    connection.Open();

                    // Actualizamoc la tabla TOTAL USUARIOS
                    string updateQueryTUS = $"UPDATE DatosIndirectos " +
                    $"SET[TOTAL USUARIOS] = (" +
                    $"    ISNULL([USUARIOS AGUA], 0) +  " +
                    $"    ISNULL([USUARIOS CLOACA], 0) + " +
                    $"    ISNULL([USUARIOS AG Y CL], 0) " +
                    $"); ";

                    // Actualizamoc la tabla % Cobertura por Conexion (Agua)]
                    string updateQueryCPC = $"UPDATE DatosIndirectos " +
                    $"SET[% COBERTURA POR CONEXION (AGUA)] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN D.POBLACION = 0 THEN NULL " +
                    $"    ELSE(((S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)] + S.[AGUA (SERVICIO MEDIDO - medidores)]) * 3.6) / NULLIF(D.POBLACION, 0)) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % Cobertura por Conexion (Cloaca)]
                    string updateQueryCPCC = $"UPDATE DatosIndirectos " +
                    $"SET[% COBERTURA POR CONEXION (CLOACA)] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN D.POBLACION = 0 THEN NULL " +
                    $"    ELSE(((S.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)]) * 3.6) / NULLIF(D.POBLACION, 0)) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % Cobertura por Conexion (Ag y CL)]
                    string updateQueryCPCAyC = $"UPDATE DatosIndirectos " +
                    $"SET[% COBERTURA POR CONEXION (AyC)] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN D.POBLACION = 0 THEN NULL " +
                    $"    ELSE(((S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)] + S.[AGUA Y CLOACA (SERVICIO MEDIDO - medidores)]) * 3.6) / NULLIF(D.POBLACION, 0)) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % COBERTURA POR USUARIO (AGUA)
                    string updateQueryCPUA = $"UPDATE DatosIndirectos " +
                    $"SET[% COBERTURA POR USUARIO (AGUA)] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN D.POBLACION = 0 THEN NULL " +
                    $"    ELSE(((I.[USUARIOS AGUA]) * 3.6) / NULLIF(D.POBLACION, 0)) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"INNER JOIN DatosIndirectos I ON DI.ID_DatosIndirectos = I.ID " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % COBERTURA POR USUARIO (CLOACA)
                    string updateQueryCPUC = $"UPDATE DatosIndirectos " +
                    $"SET[% COBERTURA POR USUARIO (CLOACA)] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN D.POBLACION = 0 THEN NULL " +
                    $"    ELSE(((I.[USUARIOS CLOACA]) * 3.6) / NULLIF(D.POBLACION, 0)) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"INNER JOIN DatosIndirectos I ON DI.ID_DatosIndirectos = I.ID " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % COBERTURA POR USUARIO (AyC)
                    string updateQueryCPUAyC = $"UPDATE DatosIndirectos " +
                    $"SET[% COBERTURA POR USUARIO (AyC)] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN D.POBLACION = 0 THEN NULL " +
                    $"    ELSE(((I.[USUARIOS AG Y CL]) * 3.6) / NULLIF(D.POBLACION, 0)) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"INNER JOIN DatosIndirectos I ON DI.ID_DatosIndirectos = I.ID " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % MICROMEDICION
                    string updateQueryMICRO = $"UPDATE DatosIndirectos " +
                    $"SET[% MICROMEDICION] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN (ISNULL(NULLIF(S.[TOTAL SERVICIO MEDIDO], 0), 0) + ISNULL(NULLIF(S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)], 0), 0) + ISNULL(NULLIF(S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)], 0), 0)) = 0 THEN NULL " +
                    $"    ELSE ((CAST(ISNULL(NULLIF(S.[TOTAL SERVICIO MEDIDO],0),0) AS DECIMAL(18,10)) / (ISNULL(NULLIF(S.[TOTAL SERVICIO MEDIDO], 0),0) + ISNULL(NULLIF(S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)], 0),0) + ISNULL(NULLIF(S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)], 0),0)))) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"INNER JOIN DatosIndirectos I ON DI.ID_DatosIndirectos = I.ID " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    // Actualizamoc la tabla % de empleados x 1000conex
                    string updateQueryEMP = $"UPDATE DatosIndirectos " +
                    $"SET[% de empleados x 1000conex] = (" +
                    $"SELECT " +
                    $"FORMAT(" +
                    $"    CASE " +
                    $"    WHEN (ISNULL(NULLIF(S.[TOTAL SERVICIO MEDIDO], 0), 0) + ISNULL(NULLIF(S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)], 0), 0) + ISNULL(NULLIF(S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)], 0), 0) + ISNULL(NULLIF(S.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)], 0), 0)) = 0 THEN NULL " +
                    $"    ELSE " +
                    $"(((CAST(ISNULL(NULLIF(S.[DOTACION ACTIVA],0),0) AS DECIMAL(18,10))*1000)/7)/" +
                    $"(ISNULL(NULLIF(S.[CLOACA (CON CONEXION - SERVICIO NO MEDIDO)], 0), 0) " +
                    $"+ ISNULL(NULLIF(S.[AGUA (CON CONEXION - SERVICIO NO MEDIDO)],0),0) " +
                    $"+ ISNULL(NULLIF(S.[AG Y CL (CON CONEXION - SERVICIO NO MEDIDO)],0),0) " +
                    $"+ ISNULL(NULLIF(S.[TOTAL SERVICIO MEDIDO],0),0))) " +
                    $"    END, " +
                    $"    '0%' " +
                    $") " +
                    $"FROM DatosDeLocalidades D " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosServicio SL ON D.ID = SL.ID_DatosLocalidades " +
                    $"INNER JOIN dbo.DatosServicio S ON SL.ID_DatosServicio = S.ID " +
                    $"INNER JOIN dbo.DatosLocalidades_DatosIndirectos DI ON D.ID = DI.ID_DatosLocalidades " +
                    $"INNER JOIN DatosIndirectos I ON DI.ID_DatosIndirectos = I.ID " +
                    $"WHERE DI.ID_DatosIndirectos = DatosIndirectos.ID" +
                    $")";

                    SqlCommand commandTUS = new SqlCommand(updateQueryTUS, connection);
                    SqlCommand commandCPC = new SqlCommand(updateQueryCPC, connection);
                    SqlCommand commandCPCC = new SqlCommand (updateQueryCPCC, connection);
                    SqlCommand commandCPCAyC = new SqlCommand (updateQueryCPCAyC, connection);
                    SqlCommand commandCPUA = new SqlCommand (updateQueryCPUA, connection);
                    SqlCommand commandCPUC = new SqlCommand(updateQueryCPUC, connection);
                    SqlCommand commandCPUAyC = new SqlCommand(updateQueryCPUAyC, connection);
                    SqlCommand commandMICRO = new SqlCommand(updateQueryMICRO, connection);
                    SqlCommand commandEMP = new SqlCommand(updateQueryEMP, connection);

                    commandTUS.ExecuteNonQuery();
                    commandCPC.ExecuteNonQuery();
                    commandCPCC.ExecuteNonQuery();
                    commandCPCAyC.ExecuteNonQuery();
                    commandCPUA.ExecuteNonQuery();
                    commandCPUC.ExecuteNonQuery();
                    commandCPUAyC.ExecuteNonQuery();
                    commandMICRO.ExecuteNonQuery();
                    commandEMP.ExecuteNonQuery();
                    connection.Close();
                }
                    
                }catch (Exception ex)
            {
                MessageBox.Show($"Error al actualizar las formulas: {ex.Message}", "Mensaje", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
