﻿namespace CRMagazine
{
    partial class frmEstoqueEntrada
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmEstoqueEntrada));
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnSelecionar = new System.Windows.Forms.Button();
            this.lblContador = new System.Windows.Forms.Label();
            this.lblUsuario = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.txtFeedBack = new System.Windows.Forms.TextBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.Location = new System.Drawing.Point(58, 420);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(16, 24);
            this.lblStatus.TabIndex = 191;
            this.lblStatus.Text = ".";
            // 
            // btnSelecionar
            // 
            this.btnSelecionar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelecionar.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelecionar.Location = new System.Drawing.Point(484, 336);
            this.btnSelecionar.Name = "btnSelecionar";
            this.btnSelecionar.Size = new System.Drawing.Size(130, 59);
            this.btnSelecionar.TabIndex = 190;
            this.btnSelecionar.Text = "SELECIONAR ARQUIVO";
            this.btnSelecionar.UseVisualStyleBackColor = true;
            this.btnSelecionar.Click += new System.EventHandler(this.btnSelecionar_Click);
            // 
            // lblContador
            // 
            this.lblContador.AutoSize = true;
            this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblContador.Location = new System.Drawing.Point(223, 54);
            this.lblContador.Name = "lblContador";
            this.lblContador.Size = new System.Drawing.Size(16, 24);
            this.lblContador.TabIndex = 189;
            this.lblContador.Text = ".";
            // 
            // lblUsuario
            // 
            this.lblUsuario.AutoSize = true;
            this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUsuario.Location = new System.Drawing.Point(223, 15);
            this.lblUsuario.Name = "lblUsuario";
            this.lblUsuario.Size = new System.Drawing.Size(16, 24);
            this.lblUsuario.TabIndex = 188;
            this.lblUsuario.Text = ".";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(79, 103);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(512, 59);
            this.label5.TabIndex = 186;
            this.label5.Text = "ENTRADA DE MATERIAIS";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::CRMagazine.Properties.Resources.B1;
            this.pictureBox1.Location = new System.Drawing.Point(12, 5);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(205, 95);
            this.pictureBox1.TabIndex = 187;
            this.pictureBox1.TabStop = false;
            // 
            // txtFeedBack
            // 
            this.txtFeedBack.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFeedBack.Location = new System.Drawing.Point(37, 165);
            this.txtFeedBack.Multiline = true;
            this.txtFeedBack.Name = "txtFeedBack";
            this.txtFeedBack.ReadOnly = true;
            this.txtFeedBack.Size = new System.Drawing.Size(428, 158);
            this.txtFeedBack.TabIndex = 192;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::CRMagazine.Properties.Resources.lupa_24x24;
            this.pictureBox2.Location = new System.Drawing.Point(626, 15);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(25, 28);
            this.pictureBox2.TabIndex = 267;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.pictureBox2_MouseDoubleClick);
            this.pictureBox2.MouseEnter += new System.EventHandler(this.pictureBox2_MouseEnter);
            // 
            // frmEstoqueEntrada
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(663, 463);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.txtFeedBack);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnSelecionar);
            this.Controls.Add(this.lblContador);
            this.Controls.Add(this.lblUsuario);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label5);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmEstoqueEntrada";
            this.Text = "ENTRADA ESTOQUE";
            this.Load += new System.EventHandler(this.frmEstoqueEntrada_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnSelecionar;
        private System.Windows.Forms.Label lblContador;
        private System.Windows.Forms.Label lblUsuario;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtFeedBack;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}