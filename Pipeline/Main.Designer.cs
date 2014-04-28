namespace Pipeline
{
    partial class Main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.gpAnterior = new System.Windows.Forms.GroupBox();
            this.txtExcelAnterior = new System.Windows.Forms.TextBox();
            this.btnAbrirAnterior = new System.Windows.Forms.Button();
            this.gpActual = new System.Windows.Forms.GroupBox();
            this.txtExcelActual = new System.Windows.Forms.TextBox();
            this.btnAbrirActual = new System.Windows.Forms.Button();
            this.btnEjecutar = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.gpAnterior.SuspendLayout();
            this.gpActual.SuspendLayout();
            this.SuspendLayout();
            // 
            // gpAnterior
            // 
            this.gpAnterior.Controls.Add(this.txtExcelAnterior);
            this.gpAnterior.Controls.Add(this.btnAbrirAnterior);
            this.gpAnterior.Location = new System.Drawing.Point(24, 22);
            this.gpAnterior.Name = "gpAnterior";
            this.gpAnterior.Size = new System.Drawing.Size(523, 54);
            this.gpAnterior.TabIndex = 0;
            this.gpAnterior.TabStop = false;
            this.gpAnterior.Text = "Excel Anterior";
            // 
            // txtExcelAnterior
            // 
            this.txtExcelAnterior.Location = new System.Drawing.Point(6, 21);
            this.txtExcelAnterior.Name = "txtExcelAnterior";
            this.txtExcelAnterior.Size = new System.Drawing.Size(411, 20);
            this.txtExcelAnterior.TabIndex = 2;
            // 
            // btnAbrirAnterior
            // 
            this.btnAbrirAnterior.Location = new System.Drawing.Point(423, 19);
            this.btnAbrirAnterior.Name = "btnAbrirAnterior";
            this.btnAbrirAnterior.Size = new System.Drawing.Size(75, 23);
            this.btnAbrirAnterior.TabIndex = 0;
            this.btnAbrirAnterior.Text = "abrir";
            this.btnAbrirAnterior.UseVisualStyleBackColor = true;
            this.btnAbrirAnterior.Click += new System.EventHandler(this.btnAbrirAnterior_Click);
            // 
            // gpActual
            // 
            this.gpActual.Controls.Add(this.txtExcelActual);
            this.gpActual.Controls.Add(this.btnAbrirActual);
            this.gpActual.Location = new System.Drawing.Point(22, 103);
            this.gpActual.Name = "gpActual";
            this.gpActual.Size = new System.Drawing.Size(525, 54);
            this.gpActual.TabIndex = 1;
            this.gpActual.TabStop = false;
            this.gpActual.Text = "Excel Actual";
            // 
            // txtExcelActual
            // 
            this.txtExcelActual.Location = new System.Drawing.Point(8, 22);
            this.txtExcelActual.Name = "txtExcelActual";
            this.txtExcelActual.Size = new System.Drawing.Size(411, 20);
            this.txtExcelActual.TabIndex = 1;
            // 
            // btnAbrirActual
            // 
            this.btnAbrirActual.Location = new System.Drawing.Point(425, 22);
            this.btnAbrirActual.Name = "btnAbrirActual";
            this.btnAbrirActual.Size = new System.Drawing.Size(75, 23);
            this.btnAbrirActual.TabIndex = 0;
            this.btnAbrirActual.Text = "abrir";
            this.btnAbrirActual.UseVisualStyleBackColor = true;
            this.btnAbrirActual.Click += new System.EventHandler(this.btnAbrirActual_Click);
            // 
            // btnEjecutar
            // 
            this.btnEjecutar.Location = new System.Drawing.Point(165, 185);
            this.btnEjecutar.Name = "btnEjecutar";
            this.btnEjecutar.Size = new System.Drawing.Size(249, 41);
            this.btnEjecutar.TabIndex = 2;
            this.btnEjecutar.Text = "Ejecutar";
            this.btnEjecutar.UseVisualStyleBackColor = true;
            this.btnEjecutar.Click += new System.EventHandler(this.btnEjecutar_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(24, 185);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(523, 23);
            this.progressBar.TabIndex = 3;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(575, 238);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnEjecutar);
            this.Controls.Add(this.gpActual);
            this.Controls.Add(this.gpAnterior);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Main";
            this.Text = "Pipeline";
            this.gpAnterior.ResumeLayout(false);
            this.gpAnterior.PerformLayout();
            this.gpActual.ResumeLayout(false);
            this.gpActual.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gpAnterior;
        private System.Windows.Forms.GroupBox gpActual;
        private System.Windows.Forms.Button btnEjecutar;
        private System.Windows.Forms.TextBox txtExcelAnterior;
        private System.Windows.Forms.Button btnAbrirAnterior;
        private System.Windows.Forms.TextBox txtExcelActual;
        private System.Windows.Forms.Button btnAbrirActual;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

