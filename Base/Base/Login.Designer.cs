namespace Base
{
    partial class Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.textUsuario = new System.Windows.Forms.TextBox();
            this.textPass = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.picMostrar = new System.Windows.Forms.PictureBox();
            this.picOcultar = new System.Windows.Forms.PictureBox();
            this.btnIniciarSesion = new FontAwesome.Sharp.IconButton();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnCerrar = new FontAwesome.Sharp.IconButton();
            this.btnMax = new FontAwesome.Sharp.IconButton();
            this.btnMin = new FontAwesome.Sharp.IconButton();
            ((System.ComponentModel.ISupportInitialize)(this.picMostrar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picOcultar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // textUsuario
            // 
            this.textUsuario.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.textUsuario.Location = new System.Drawing.Point(112, 66);
            this.textUsuario.Name = "textUsuario";
            this.textUsuario.Size = new System.Drawing.Size(272, 20);
            this.textUsuario.TabIndex = 4;
            // 
            // textPass
            // 
            this.textPass.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.textPass.Location = new System.Drawing.Point(112, 114);
            this.textPass.Name = "textPass";
            this.textPass.PasswordChar = '*';
            this.textPass.Size = new System.Drawing.Size(272, 20);
            this.textPass.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(46, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(43, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Usuario";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 118);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Contraseña";
            // 
            // picMostrar
            // 
            this.picMostrar.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.picMostrar.BackColor = System.Drawing.Color.White;
            this.picMostrar.Image = global::Base.Properties.Resources.eye_solid;
            this.picMostrar.Location = new System.Drawing.Point(363, 115);
            this.picMostrar.Name = "picMostrar";
            this.picMostrar.Size = new System.Drawing.Size(18, 18);
            this.picMostrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picMostrar.TabIndex = 10;
            this.picMostrar.TabStop = false;
            this.picMostrar.Click += new System.EventHandler(this.picMostrar_Click);
            // 
            // picOcultar
            // 
            this.picOcultar.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.picOcultar.BackColor = System.Drawing.Color.White;
            this.picOcultar.Image = global::Base.Properties.Resources.eye_slash_solid;
            this.picOcultar.Location = new System.Drawing.Point(363, 115);
            this.picOcultar.Name = "picOcultar";
            this.picOcultar.Size = new System.Drawing.Size(18, 18);
            this.picOcultar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picOcultar.TabIndex = 9;
            this.picOcultar.TabStop = false;
            this.picOcultar.Click += new System.EventHandler(this.picOcultar_Click);
            // 
            // btnIniciarSesion
            // 
            this.btnIniciarSesion.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnIniciarSesion.IconChar = FontAwesome.Sharp.IconChar.None;
            this.btnIniciarSesion.IconColor = System.Drawing.Color.Black;
            this.btnIniciarSesion.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.btnIniciarSesion.Location = new System.Drawing.Point(186, 156);
            this.btnIniciarSesion.Name = "btnIniciarSesion";
            this.btnIniciarSesion.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btnIniciarSesion.Size = new System.Drawing.Size(112, 23);
            this.btnIniciarSesion.TabIndex = 6;
            this.btnIniciarSesion.Text = "Iniciar Sesión";
            this.btnIniciarSesion.UseVisualStyleBackColor = true;
            this.btnIniciarSesion.Click += new System.EventHandler(this.btnIniciarSesion_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Base.Properties.Resources.LogoSameep;
            this.pictureBox1.Location = new System.Drawing.Point(7, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(79, 50);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // btnCerrar
            // 
            this.btnCerrar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCerrar.FlatAppearance.BorderSize = 0;
            this.btnCerrar.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnCerrar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Red;
            this.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCerrar.IconChar = FontAwesome.Sharp.IconChar.Xmark;
            this.btnCerrar.IconColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnCerrar.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.btnCerrar.IconSize = 25;
            this.btnCerrar.Location = new System.Drawing.Point(451, 4);
            this.btnCerrar.Name = "btnCerrar";
            this.btnCerrar.Size = new System.Drawing.Size(35, 25);
            this.btnCerrar.TabIndex = 2;
            this.btnCerrar.UseVisualStyleBackColor = true;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);
            // 
            // btnMax
            // 
            this.btnMax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMax.FlatAppearance.BorderSize = 0;
            this.btnMax.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.btnMax.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.btnMax.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMax.IconChar = FontAwesome.Sharp.IconChar.Expand;
            this.btnMax.IconColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnMax.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.btnMax.IconSize = 25;
            this.btnMax.Location = new System.Drawing.Point(416, 4);
            this.btnMax.Name = "btnMax";
            this.btnMax.Size = new System.Drawing.Size(35, 25);
            this.btnMax.TabIndex = 1;
            this.btnMax.UseVisualStyleBackColor = true;
            this.btnMax.Click += new System.EventHandler(this.btnMax_Click);
            // 
            // btnMin
            // 
            this.btnMin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMin.FlatAppearance.BorderSize = 0;
            this.btnMin.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.btnMin.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Silver;
            this.btnMin.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnMin.IconChar = FontAwesome.Sharp.IconChar.Minus;
            this.btnMin.IconColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnMin.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.btnMin.IconSize = 25;
            this.btnMin.Location = new System.Drawing.Point(381, 4);
            this.btnMin.Name = "btnMin";
            this.btnMin.Size = new System.Drawing.Size(35, 25);
            this.btnMin.TabIndex = 0;
            this.btnMin.UseVisualStyleBackColor = true;
            this.btnMin.Click += new System.EventHandler(this.btnMin_Click);
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
            this.ClientSize = new System.Drawing.Size(490, 205);
            this.Controls.Add(this.picMostrar);
            this.Controls.Add(this.picOcultar);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnIniciarSesion);
            this.Controls.Add(this.textPass);
            this.Controls.Add(this.textUsuario);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnCerrar);
            this.Controls.Add(this.btnMax);
            this.Controls.Add(this.btnMin);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login";
            ((System.ComponentModel.ISupportInitialize)(this.picMostrar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picOcultar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private FontAwesome.Sharp.IconButton btnMin;
        private FontAwesome.Sharp.IconButton btnMax;
        private FontAwesome.Sharp.IconButton btnCerrar;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textUsuario;
        private System.Windows.Forms.TextBox textPass;
        private FontAwesome.Sharp.IconButton btnIniciarSesion;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox picOcultar;
        private System.Windows.Forms.PictureBox picMostrar;
    }
}