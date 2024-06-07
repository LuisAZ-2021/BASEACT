namespace Base
{
    partial class Inicio
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Inicio));
            this.menuPrincipal = new System.Windows.Forms.MenuStrip();
            this.menuGeneral = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarGeneral = new FontAwesome.Sharp.IconMenuItem();
            this.menuLocalidad = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarLocalidad = new FontAwesome.Sharp.IconMenuItem();
            this.submenuCargarLocalidad = new FontAwesome.Sharp.IconMenuItem();
            this.submenuActualizarLocalidad = new FontAwesome.Sharp.IconMenuItem();
            this.menuFuente = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarFuente = new FontAwesome.Sharp.IconMenuItem();
            this.submenuActualizarFuente = new FontAwesome.Sharp.IconMenuItem();
            this.menuServicio = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarServicio = new FontAwesome.Sharp.IconMenuItem();
            this.submenuActualizarServicio = new FontAwesome.Sharp.IconMenuItem();
            this.menuFinanciero = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarFinanciero = new FontAwesome.Sharp.IconMenuItem();
            this.submenuActualizarFinanciero = new FontAwesome.Sharp.IconMenuItem();
            this.menuIndirecto = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarIndirecto = new FontAwesome.Sharp.IconMenuItem();
            this.submenuActualizarIndirecto = new FontAwesome.Sharp.IconMenuItem();
            this.menuUsuarios = new FontAwesome.Sharp.IconMenuItem();
            this.submenuMostrarUsuarios = new FontAwesome.Sharp.IconMenuItem();
            this.submenuCargarUsuarios = new FontAwesome.Sharp.IconMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.lblBienvenida = new System.Windows.Forms.Label();
            this.lblRol = new System.Windows.Forms.Label();
            this.menuEncabezado = new System.Windows.Forms.MenuStrip();
            this.contenedor = new System.Windows.Forms.Panel();
            this.menuPrincipal.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuPrincipal
            // 
            this.menuPrincipal.AutoSize = false;
            this.menuPrincipal.BackColor = System.Drawing.Color.LightBlue;
            this.menuPrincipal.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuGeneral,
            this.menuLocalidad,
            this.menuFuente,
            this.menuServicio,
            this.menuFinanciero,
            this.menuIndirecto,
            this.menuUsuarios});
            this.menuPrincipal.Location = new System.Drawing.Point(0, 52);
            this.menuPrincipal.Name = "menuPrincipal";
            this.menuPrincipal.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.menuPrincipal.Size = new System.Drawing.Size(941, 51);
            this.menuPrincipal.TabIndex = 0;
            this.menuPrincipal.Text = "menuEncabezado";
            // 
            // menuGeneral
            // 
            this.menuGeneral.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarGeneral});
            this.menuGeneral.IconChar = FontAwesome.Sharp.IconChar.House;
            this.menuGeneral.IconColor = System.Drawing.Color.Black;
            this.menuGeneral.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuGeneral.IconSize = 30;
            this.menuGeneral.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuGeneral.Name = "menuGeneral";
            this.menuGeneral.Size = new System.Drawing.Size(59, 47);
            this.menuGeneral.Text = "General";
            this.menuGeneral.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarGeneral
            // 
            this.submenuMostrarGeneral.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarGeneral.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarGeneral.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarGeneral.Name = "submenuMostrarGeneral";
            this.submenuMostrarGeneral.Size = new System.Drawing.Size(153, 22);
            this.submenuMostrarGeneral.Text = "Detalle General";
            this.submenuMostrarGeneral.Click += new System.EventHandler(this.submenuMostrarGeneral_Click);
            // 
            // menuLocalidad
            // 
            this.menuLocalidad.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarLocalidad,
            this.submenuCargarLocalidad,
            this.submenuActualizarLocalidad});
            this.menuLocalidad.IconChar = FontAwesome.Sharp.IconChar.MapLocation;
            this.menuLocalidad.IconColor = System.Drawing.Color.Black;
            this.menuLocalidad.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuLocalidad.IconSize = 30;
            this.menuLocalidad.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuLocalidad.Name = "menuLocalidad";
            this.menuLocalidad.Size = new System.Drawing.Size(70, 47);
            this.menuLocalidad.Text = "Localidad";
            this.menuLocalidad.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarLocalidad
            // 
            this.submenuMostrarLocalidad.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarLocalidad.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarLocalidad.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarLocalidad.Name = "submenuMostrarLocalidad";
            this.submenuMostrarLocalidad.Size = new System.Drawing.Size(180, 22);
            this.submenuMostrarLocalidad.Text = "Detalle de Localidad";
            this.submenuMostrarLocalidad.Click += new System.EventHandler(this.submenuMostrarLocalidad_Click);
            // 
            // submenuCargarLocalidad
            // 
            this.submenuCargarLocalidad.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuCargarLocalidad.IconColor = System.Drawing.Color.Black;
            this.submenuCargarLocalidad.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuCargarLocalidad.Name = "submenuCargarLocalidad";
            this.submenuCargarLocalidad.Size = new System.Drawing.Size(180, 22);
            this.submenuCargarLocalidad.Text = "Cargar Localidad";
            this.submenuCargarLocalidad.Click += new System.EventHandler(this.submenuCargarLocalidad_Click);
            // 
            // submenuActualizarLocalidad
            // 
            this.submenuActualizarLocalidad.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuActualizarLocalidad.IconColor = System.Drawing.Color.Black;
            this.submenuActualizarLocalidad.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuActualizarLocalidad.Name = "submenuActualizarLocalidad";
            this.submenuActualizarLocalidad.Size = new System.Drawing.Size(180, 22);
            this.submenuActualizarLocalidad.Text = "Actualizar Localidad";
            this.submenuActualizarLocalidad.Click += new System.EventHandler(this.submenuActualizarLocalidad_Click);
            // 
            // menuFuente
            // 
            this.menuFuente.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarFuente,
            this.submenuActualizarFuente});
            this.menuFuente.IconChar = FontAwesome.Sharp.IconChar.Water;
            this.menuFuente.IconColor = System.Drawing.Color.Black;
            this.menuFuente.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuFuente.IconSize = 30;
            this.menuFuente.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuFuente.Name = "menuFuente";
            this.menuFuente.Size = new System.Drawing.Size(55, 47);
            this.menuFuente.Text = "Fuente";
            this.menuFuente.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarFuente
            // 
            this.submenuMostrarFuente.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarFuente.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarFuente.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarFuente.Name = "submenuMostrarFuente";
            this.submenuMostrarFuente.Size = new System.Drawing.Size(165, 22);
            this.submenuMostrarFuente.Text = "Detalle de Fuente";
            this.submenuMostrarFuente.Click += new System.EventHandler(this.submenuMostrarFuente_Click);
            // 
            // submenuActualizarFuente
            // 
            this.submenuActualizarFuente.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuActualizarFuente.IconColor = System.Drawing.Color.Black;
            this.submenuActualizarFuente.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuActualizarFuente.Name = "submenuActualizarFuente";
            this.submenuActualizarFuente.Size = new System.Drawing.Size(165, 22);
            this.submenuActualizarFuente.Text = "Actualizar Fuente";
            this.submenuActualizarFuente.Click += new System.EventHandler(this.submenuActualizarFuente_Click);
            // 
            // menuServicio
            // 
            this.menuServicio.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarServicio,
            this.submenuActualizarServicio});
            this.menuServicio.IconChar = FontAwesome.Sharp.IconChar.HandHoldingDroplet;
            this.menuServicio.IconColor = System.Drawing.Color.Black;
            this.menuServicio.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuServicio.IconSize = 30;
            this.menuServicio.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuServicio.Name = "menuServicio";
            this.menuServicio.Size = new System.Drawing.Size(60, 47);
            this.menuServicio.Text = "Servicio";
            this.menuServicio.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarServicio
            // 
            this.submenuMostrarServicio.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarServicio.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarServicio.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarServicio.Name = "submenuMostrarServicio";
            this.submenuMostrarServicio.Size = new System.Drawing.Size(173, 22);
            this.submenuMostrarServicio.Text = "Detalle del Servicio";
            this.submenuMostrarServicio.Click += new System.EventHandler(this.submenuMostrarServicio_Click);
            // 
            // submenuActualizarServicio
            // 
            this.submenuActualizarServicio.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuActualizarServicio.IconColor = System.Drawing.Color.Black;
            this.submenuActualizarServicio.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuActualizarServicio.Name = "submenuActualizarServicio";
            this.submenuActualizarServicio.Size = new System.Drawing.Size(173, 22);
            this.submenuActualizarServicio.Text = "Actualizar Servicio";
            this.submenuActualizarServicio.Click += new System.EventHandler(this.submenuActualizarServicio_Click);
            // 
            // menuFinanciero
            // 
            this.menuFinanciero.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarFinanciero,
            this.submenuActualizarFinanciero});
            this.menuFinanciero.IconChar = FontAwesome.Sharp.IconChar.HandHoldingUsd;
            this.menuFinanciero.IconColor = System.Drawing.Color.Black;
            this.menuFinanciero.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuFinanciero.IconSize = 30;
            this.menuFinanciero.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuFinanciero.Name = "menuFinanciero";
            this.menuFinanciero.Size = new System.Drawing.Size(112, 47);
            this.menuFinanciero.Text = "Datos Financieros";
            this.menuFinanciero.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarFinanciero
            // 
            this.submenuMostrarFinanciero.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarFinanciero.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarFinanciero.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarFinanciero.Name = "submenuMostrarFinanciero";
            this.submenuMostrarFinanciero.Size = new System.Drawing.Size(222, 22);
            this.submenuMostrarFinanciero.Text = "Detalle Financiero";
            this.submenuMostrarFinanciero.Click += new System.EventHandler(this.submenuMostrarFinanciero_Click);
            // 
            // submenuActualizarFinanciero
            // 
            this.submenuActualizarFinanciero.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuActualizarFinanciero.IconColor = System.Drawing.Color.Black;
            this.submenuActualizarFinanciero.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuActualizarFinanciero.Name = "submenuActualizarFinanciero";
            this.submenuActualizarFinanciero.Size = new System.Drawing.Size(222, 22);
            this.submenuActualizarFinanciero.Text = "Actualizar Datos Financieros";
            this.submenuActualizarFinanciero.Click += new System.EventHandler(this.submenuActualizarFinanciero_Click);
            // 
            // menuIndirecto
            // 
            this.menuIndirecto.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarIndirecto,
            this.submenuActualizarIndirecto});
            this.menuIndirecto.IconChar = FontAwesome.Sharp.IconChar.Percentage;
            this.menuIndirecto.IconColor = System.Drawing.Color.Black;
            this.menuIndirecto.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuIndirecto.IconSize = 30;
            this.menuIndirecto.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuIndirecto.Name = "menuIndirecto";
            this.menuIndirecto.Size = new System.Drawing.Size(104, 47);
            this.menuIndirecto.Text = "Datos Indirectos";
            this.menuIndirecto.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarIndirecto
            // 
            this.submenuMostrarIndirecto.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarIndirecto.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarIndirecto.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarIndirecto.Name = "submenuMostrarIndirecto";
            this.submenuMostrarIndirecto.Size = new System.Drawing.Size(214, 22);
            this.submenuMostrarIndirecto.Text = "Detalle Datos Indirectos";
            this.submenuMostrarIndirecto.Click += new System.EventHandler(this.submenuMostrarIndirecto_Click);
            // 
            // submenuActualizarIndirecto
            // 
            this.submenuActualizarIndirecto.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuActualizarIndirecto.IconColor = System.Drawing.Color.Black;
            this.submenuActualizarIndirecto.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuActualizarIndirecto.Name = "submenuActualizarIndirecto";
            this.submenuActualizarIndirecto.Size = new System.Drawing.Size(214, 22);
            this.submenuActualizarIndirecto.Text = "Actualizar Datos Indirectos";
            this.submenuActualizarIndirecto.Click += new System.EventHandler(this.submenuActualizarIndirecto_Click);
            // 
            // menuUsuarios
            // 
            this.menuUsuarios.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.submenuMostrarUsuarios,
            this.submenuCargarUsuarios});
            this.menuUsuarios.IconChar = FontAwesome.Sharp.IconChar.Users;
            this.menuUsuarios.IconColor = System.Drawing.Color.Black;
            this.menuUsuarios.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.menuUsuarios.IconSize = 30;
            this.menuUsuarios.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.menuUsuarios.Name = "menuUsuarios";
            this.menuUsuarios.Size = new System.Drawing.Size(64, 47);
            this.menuUsuarios.Text = "Usuarios";
            this.menuUsuarios.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // submenuMostrarUsuarios
            // 
            this.submenuMostrarUsuarios.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuMostrarUsuarios.IconColor = System.Drawing.Color.Black;
            this.submenuMostrarUsuarios.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuMostrarUsuarios.Name = "submenuMostrarUsuarios";
            this.submenuMostrarUsuarios.Size = new System.Drawing.Size(180, 22);
            this.submenuMostrarUsuarios.Text = "Mostrar Usuarios";
            this.submenuMostrarUsuarios.Click += new System.EventHandler(this.submenuMostrarUsuarios_Click);
            // 
            // submenuCargarUsuarios
            // 
            this.submenuCargarUsuarios.IconChar = FontAwesome.Sharp.IconChar.None;
            this.submenuCargarUsuarios.IconColor = System.Drawing.Color.Black;
            this.submenuCargarUsuarios.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.submenuCargarUsuarios.Name = "submenuCargarUsuarios";
            this.submenuCargarUsuarios.Size = new System.Drawing.Size(180, 22);
            this.submenuCargarUsuarios.Text = "Cargar Usuarios";
            this.submenuCargarUsuarios.Click += new System.EventHandler(this.submenuCargarUsuarios_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.SteelBlue;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(16, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(218, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Base de Datos DGPE";
            // 
            // lblBienvenida
            // 
            this.lblBienvenida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblBienvenida.AutoSize = true;
            this.lblBienvenida.BackColor = System.Drawing.Color.SteelBlue;
            this.lblBienvenida.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBienvenida.ForeColor = System.Drawing.Color.White;
            this.lblBienvenida.Location = new System.Drawing.Point(834, 7);
            this.lblBienvenida.Name = "lblBienvenida";
            this.lblBienvenida.Size = new System.Drawing.Size(68, 15);
            this.lblBienvenida.TabIndex = 2;
            this.lblBienvenida.Text = "Bienvenida";
            // 
            // lblRol
            // 
            this.lblRol.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblRol.AutoSize = true;
            this.lblRol.BackColor = System.Drawing.Color.SteelBlue;
            this.lblRol.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRol.ForeColor = System.Drawing.Color.White;
            this.lblRol.Location = new System.Drawing.Point(841, 27);
            this.lblRol.Name = "lblRol";
            this.lblRol.Size = new System.Drawing.Size(26, 15);
            this.lblRol.TabIndex = 3;
            this.lblRol.Text = "Rol";
            // 
            // menuEncabezado
            // 
            this.menuEncabezado.AutoSize = false;
            this.menuEncabezado.BackColor = System.Drawing.Color.SteelBlue;
            this.menuEncabezado.Location = new System.Drawing.Point(0, 0);
            this.menuEncabezado.Name = "menuEncabezado";
            this.menuEncabezado.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.menuEncabezado.Size = new System.Drawing.Size(941, 52);
            this.menuEncabezado.TabIndex = 4;
            this.menuEncabezado.Text = "menuStrip1";
            // 
            // contenedor
            // 
            this.contenedor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contenedor.Location = new System.Drawing.Point(0, 103);
            this.contenedor.Name = "contenedor";
            this.contenedor.Size = new System.Drawing.Size(941, 376);
            this.contenedor.TabIndex = 5;
            // 
            // Inicio
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(941, 479);
            this.Controls.Add(this.contenedor);
            this.Controls.Add(this.lblRol);
            this.Controls.Add(this.lblBienvenida);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.menuPrincipal);
            this.Controls.Add(this.menuEncabezado);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuPrincipal;
            this.Name = "Inicio";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Inicio";
            this.Load += new System.EventHandler(this.Inicio_Load);
            this.menuPrincipal.ResumeLayout(false);
            this.menuPrincipal.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuPrincipal;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblBienvenida;
        private System.Windows.Forms.Label lblRol;
        private System.Windows.Forms.MenuStrip menuEncabezado;
        private FontAwesome.Sharp.IconMenuItem menuLocalidad;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarLocalidad;
        private FontAwesome.Sharp.IconMenuItem submenuCargarLocalidad;
        private FontAwesome.Sharp.IconMenuItem submenuActualizarLocalidad;
        private FontAwesome.Sharp.IconMenuItem menuUsuarios;
        private System.Windows.Forms.Panel contenedor;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarUsuarios;
        private FontAwesome.Sharp.IconMenuItem submenuCargarUsuarios;
        private FontAwesome.Sharp.IconMenuItem menuGeneral;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarGeneral;
        private FontAwesome.Sharp.IconMenuItem menuFuente;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarFuente;
        private FontAwesome.Sharp.IconMenuItem submenuActualizarFuente;
        private FontAwesome.Sharp.IconMenuItem menuServicio;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarServicio;
        private FontAwesome.Sharp.IconMenuItem submenuActualizarServicio;
        private FontAwesome.Sharp.IconMenuItem menuFinanciero;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarFinanciero;
        private FontAwesome.Sharp.IconMenuItem submenuActualizarFinanciero;
        private FontAwesome.Sharp.IconMenuItem menuIndirecto;
        private FontAwesome.Sharp.IconMenuItem submenuMostrarIndirecto;
        private FontAwesome.Sharp.IconMenuItem submenuActualizarIndirecto;
    }
}

