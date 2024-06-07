namespace Base
{
    partial class EditCellForm
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
            this.Valor1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textValor = new System.Windows.Forms.TextBox();
            this.textComentario = new System.Windows.Forms.TextBox();
            this.btnGuardar = new FontAwesome.Sharp.IconButton();
            this.comboBoxLocalidad = new System.Windows.Forms.ComboBox();
            this.Valor2 = new System.Windows.Forms.Label();
            this.comboBoxZona = new System.Windows.Forms.ComboBox();
            this.valorZona = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Valor1
            // 
            this.Valor1.AutoSize = true;
            this.Valor1.Location = new System.Drawing.Point(59, 41);
            this.Valor1.Name = "Valor1";
            this.Valor1.Size = new System.Drawing.Size(31, 13);
            this.Valor1.TabIndex = 0;
            this.Valor1.Text = "Valor";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(59, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Fuente";
            // 
            // textValor
            // 
            this.textValor.Location = new System.Drawing.Point(153, 38);
            this.textValor.Name = "textValor";
            this.textValor.Size = new System.Drawing.Size(193, 20);
            this.textValor.TabIndex = 2;
            // 
            // textComentario
            // 
            this.textComentario.Location = new System.Drawing.Point(153, 105);
            this.textComentario.Name = "textComentario";
            this.textComentario.Size = new System.Drawing.Size(193, 20);
            this.textComentario.TabIndex = 3;
            // 
            // btnGuardar
            // 
            this.btnGuardar.IconChar = FontAwesome.Sharp.IconChar.None;
            this.btnGuardar.IconColor = System.Drawing.Color.Black;
            this.btnGuardar.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.btnGuardar.Location = new System.Drawing.Point(211, 155);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 4;
            this.btnGuardar.Text = "Aceptar";
            this.btnGuardar.UseVisualStyleBackColor = true;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
            // 
            // comboBoxLocalidad
            // 
            this.comboBoxLocalidad.FormattingEnabled = true;
            this.comboBoxLocalidad.Location = new System.Drawing.Point(153, 43);
            this.comboBoxLocalidad.Name = "comboBoxLocalidad";
            this.comboBoxLocalidad.Size = new System.Drawing.Size(193, 21);
            this.comboBoxLocalidad.TabIndex = 5;
            // 
            // Valor2
            // 
            this.Valor2.AutoSize = true;
            this.Valor2.Location = new System.Drawing.Point(59, 48);
            this.Valor2.Name = "Valor2";
            this.Valor2.Size = new System.Drawing.Size(31, 13);
            this.Valor2.TabIndex = 6;
            this.Valor2.Text = "Valor";
            // 
            // comboBoxZona
            // 
            this.comboBoxZona.FormattingEnabled = true;
            this.comboBoxZona.Location = new System.Drawing.Point(153, 48);
            this.comboBoxZona.Name = "comboBoxZona";
            this.comboBoxZona.Size = new System.Drawing.Size(193, 21);
            this.comboBoxZona.TabIndex = 7;
            // 
            // valorZona
            // 
            this.valorZona.AutoSize = true;
            this.valorZona.Location = new System.Drawing.Point(59, 55);
            this.valorZona.Name = "valorZona";
            this.valorZona.Size = new System.Drawing.Size(31, 13);
            this.valorZona.TabIndex = 8;
            this.valorZona.Text = "Valor";
            // 
            // EditCellForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(427, 190);
            this.Controls.Add(this.valorZona);
            this.Controls.Add(this.comboBoxZona);
            this.Controls.Add(this.Valor2);
            this.Controls.Add(this.comboBoxLocalidad);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.textComentario);
            this.Controls.Add(this.textValor);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Valor1);
            this.Name = "EditCellForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EditCellForm";
            this.Load += new System.EventHandler(this.EditCellForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Valor1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textValor;
        private System.Windows.Forms.TextBox textComentario;
        private FontAwesome.Sharp.IconButton btnGuardar;
        private System.Windows.Forms.ComboBox comboBoxLocalidad;
        private System.Windows.Forms.Label Valor2;
        private System.Windows.Forms.ComboBox comboBoxZona;
        private System.Windows.Forms.Label valorZona;
    }
}