﻿namespace CRMagazine
{
    partial class frmInserirUsuario
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmInserirUsuario));
            this.cbxEntrada = new System.Windows.Forms.ComboBox();
            this.cbxVistoria = new System.Windows.Forms.ComboBox();
            this.cbxReparo = new System.Windows.Forms.ComboBox();
            this.cbxConsultas = new System.Windows.Forms.ComboBox();
            this.cbxExpedicao = new System.Windows.Forms.ComboBox();
            this.cbxRunin = new System.Windows.Forms.ComboBox();
            this.cbxAjuste = new System.Windows.Forms.ComboBox();
            this.cbxEmbalagem = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.cbxUsuario = new System.Windows.Forms.ComboBox();
            this.rbtNovoUsuario = new System.Windows.Forms.RadioButton();
            this.rbtUsuarioExistente = new System.Windows.Forms.RadioButton();
            this.rbtDeleteUsuario = new System.Windows.Forms.RadioButton();
            this.btnConcluir = new System.Windows.Forms.Button();
            this.txtUsuario = new System.Windows.Forms.TextBox();
            this.btnLimparTela = new System.Windows.Forms.Button();
            this.rbtResetSenha = new System.Windows.Forms.RadioButton();
            this.label9 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cbxEstoque = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.cbxAssist = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.cbxADM = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // cbxEntrada
            // 
            this.cbxEntrada.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxEntrada.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxEntrada.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxEntrada.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxEntrada.FormattingEnabled = true;
            this.cbxEntrada.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxEntrada.Location = new System.Drawing.Point(343, 238);
            this.cbxEntrada.Name = "cbxEntrada";
            this.cbxEntrada.Size = new System.Drawing.Size(63, 24);
            this.cbxEntrada.TabIndex = 33;
            this.cbxEntrada.Text = "...";
            // 
            // cbxVistoria
            // 
            this.cbxVistoria.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxVistoria.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxVistoria.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxVistoria.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxVistoria.FormattingEnabled = true;
            this.cbxVistoria.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxVistoria.Location = new System.Drawing.Point(565, 251);
            this.cbxVistoria.Name = "cbxVistoria";
            this.cbxVistoria.Size = new System.Drawing.Size(63, 24);
            this.cbxVistoria.TabIndex = 34;
            this.cbxVistoria.Text = "...";
            this.cbxVistoria.Visible = false;
            // 
            // cbxReparo
            // 
            this.cbxReparo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxReparo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxReparo.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxReparo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxReparo.FormattingEnabled = true;
            this.cbxReparo.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxReparo.Location = new System.Drawing.Point(343, 268);
            this.cbxReparo.Name = "cbxReparo";
            this.cbxReparo.Size = new System.Drawing.Size(63, 24);
            this.cbxReparo.TabIndex = 35;
            this.cbxReparo.Text = "...";
            // 
            // cbxConsultas
            // 
            this.cbxConsultas.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxConsultas.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxConsultas.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxConsultas.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxConsultas.FormattingEnabled = true;
            this.cbxConsultas.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxConsultas.Location = new System.Drawing.Point(343, 363);
            this.cbxConsultas.Name = "cbxConsultas";
            this.cbxConsultas.Size = new System.Drawing.Size(63, 24);
            this.cbxConsultas.TabIndex = 38;
            this.cbxConsultas.Text = "...";
            // 
            // cbxExpedicao
            // 
            this.cbxExpedicao.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxExpedicao.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxExpedicao.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxExpedicao.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxExpedicao.FormattingEnabled = true;
            this.cbxExpedicao.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxExpedicao.Location = new System.Drawing.Point(343, 332);
            this.cbxExpedicao.Name = "cbxExpedicao";
            this.cbxExpedicao.Size = new System.Drawing.Size(63, 24);
            this.cbxExpedicao.TabIndex = 37;
            this.cbxExpedicao.Text = "...";
            // 
            // cbxRunin
            // 
            this.cbxRunin.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxRunin.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxRunin.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxRunin.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxRunin.FormattingEnabled = true;
            this.cbxRunin.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxRunin.Location = new System.Drawing.Point(565, 281);
            this.cbxRunin.Name = "cbxRunin";
            this.cbxRunin.Size = new System.Drawing.Size(63, 24);
            this.cbxRunin.TabIndex = 36;
            this.cbxRunin.Text = "...";
            this.cbxRunin.Visible = false;
            // 
            // cbxAjuste
            // 
            this.cbxAjuste.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxAjuste.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxAjuste.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxAjuste.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxAjuste.FormattingEnabled = true;
            this.cbxAjuste.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxAjuste.Location = new System.Drawing.Point(343, 424);
            this.cbxAjuste.Name = "cbxAjuste";
            this.cbxAjuste.Size = new System.Drawing.Size(63, 24);
            this.cbxAjuste.TabIndex = 40;
            this.cbxAjuste.Text = "...";
            // 
            // cbxEmbalagem
            // 
            this.cbxEmbalagem.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxEmbalagem.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxEmbalagem.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxEmbalagem.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxEmbalagem.FormattingEnabled = true;
            this.cbxEmbalagem.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxEmbalagem.Location = new System.Drawing.Point(343, 302);
            this.cbxEmbalagem.Name = "cbxEmbalagem";
            this.cbxEmbalagem.Size = new System.Drawing.Size(63, 24);
            this.cbxEmbalagem.TabIndex = 39;
            this.cbxEmbalagem.Text = "...";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(204, 238);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 20);
            this.label1.TabIndex = 41;
            this.label1.Text = "Entrada";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(426, 251);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 20);
            this.label2.TabIndex = 42;
            this.label2.Text = "Vistoria";
            this.label2.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(204, 268);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 20);
            this.label3.TabIndex = 43;
            this.label3.Text = "Reparo";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(426, 281);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(51, 20);
            this.label4.TabIndex = 44;
            this.label4.Text = "Runin";
            this.label4.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(204, 332);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 20);
            this.label5.TabIndex = 45;
            this.label5.Text = "Expedicao";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(204, 363);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 20);
            this.label6.TabIndex = 46;
            this.label6.Text = "Consultas";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(204, 302);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(94, 20);
            this.label7.TabIndex = 47;
            this.label7.Text = "Embalagem";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(204, 424);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(54, 20);
            this.label8.TabIndex = 48;
            this.label8.Text = "Ajuste";
            // 
            // cbxUsuario
            // 
            this.cbxUsuario.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxUsuario.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxUsuario.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxUsuario.FormattingEnabled = true;
            this.cbxUsuario.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxUsuario.Location = new System.Drawing.Point(208, 187);
            this.cbxUsuario.Name = "cbxUsuario";
            this.cbxUsuario.Size = new System.Drawing.Size(198, 24);
            this.cbxUsuario.TabIndex = 50;
            this.cbxUsuario.Text = "...";
            this.cbxUsuario.Visible = false;
            this.cbxUsuario.SelectedIndexChanged += new System.EventHandler(this.cbxUsuario_SelectedIndexChanged);
            // 
            // rbtNovoUsuario
            // 
            this.rbtNovoUsuario.AutoSize = true;
            this.rbtNovoUsuario.Location = new System.Drawing.Point(26, 144);
            this.rbtNovoUsuario.Name = "rbtNovoUsuario";
            this.rbtNovoUsuario.Size = new System.Drawing.Size(90, 17);
            this.rbtNovoUsuario.TabIndex = 52;
            this.rbtNovoUsuario.TabStop = true;
            this.rbtNovoUsuario.Text = "Novo Usuário";
            this.rbtNovoUsuario.UseVisualStyleBackColor = true;
            this.rbtNovoUsuario.CheckedChanged += new System.EventHandler(this.rbtNovoUsuario_CheckedChanged);
            // 
            // rbtUsuarioExistente
            // 
            this.rbtUsuarioExistente.AutoSize = true;
            this.rbtUsuarioExistente.Location = new System.Drawing.Point(26, 167);
            this.rbtUsuarioExistente.Name = "rbtUsuarioExistente";
            this.rbtUsuarioExistente.Size = new System.Drawing.Size(111, 17);
            this.rbtUsuarioExistente.TabIndex = 53;
            this.rbtUsuarioExistente.TabStop = true;
            this.rbtUsuarioExistente.Text = "Alterar Permissões";
            this.rbtUsuarioExistente.UseVisualStyleBackColor = true;
            this.rbtUsuarioExistente.CheckedChanged += new System.EventHandler(this.rbtUsuarioExistente_CheckedChanged);
            // 
            // rbtDeleteUsuario
            // 
            this.rbtDeleteUsuario.AutoSize = true;
            this.rbtDeleteUsuario.Location = new System.Drawing.Point(26, 213);
            this.rbtDeleteUsuario.Name = "rbtDeleteUsuario";
            this.rbtDeleteUsuario.Size = new System.Drawing.Size(98, 17);
            this.rbtDeleteUsuario.TabIndex = 54;
            this.rbtDeleteUsuario.TabStop = true;
            this.rbtDeleteUsuario.Text = "Deletar Usuário";
            this.rbtDeleteUsuario.UseVisualStyleBackColor = true;
            this.rbtDeleteUsuario.CheckedChanged += new System.EventHandler(this.rbtDeleteUsuario_CheckedChanged);
            // 
            // btnConcluir
            // 
            this.btnConcluir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConcluir.Location = new System.Drawing.Point(521, 460);
            this.btnConcluir.Name = "btnConcluir";
            this.btnConcluir.Size = new System.Drawing.Size(91, 36);
            this.btnConcluir.TabIndex = 140;
            this.btnConcluir.Text = "CONCLUIR";
            this.btnConcluir.UseVisualStyleBackColor = true;
            this.btnConcluir.Click += new System.EventHandler(this.btnConcluir_Click);
            // 
            // txtUsuario
            // 
            this.txtUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtUsuario.Location = new System.Drawing.Point(208, 157);
            this.txtUsuario.Name = "txtUsuario";
            this.txtUsuario.Size = new System.Drawing.Size(198, 24);
            this.txtUsuario.TabIndex = 141;
            this.txtUsuario.Visible = false;
            // 
            // btnLimparTela
            // 
            this.btnLimparTela.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnLimparTela.Location = new System.Drawing.Point(12, 460);
            this.btnLimparTela.Name = "btnLimparTela";
            this.btnLimparTela.Size = new System.Drawing.Size(91, 36);
            this.btnLimparTela.TabIndex = 142;
            this.btnLimparTela.Text = "LIMPAR TELA";
            this.btnLimparTela.UseVisualStyleBackColor = true;
            this.btnLimparTela.Click += new System.EventHandler(this.btnLimparTela_Click);
            // 
            // rbtResetSenha
            // 
            this.rbtResetSenha.AutoSize = true;
            this.rbtResetSenha.Location = new System.Drawing.Point(26, 190);
            this.rbtResetSenha.Name = "rbtResetSenha";
            this.rbtResetSenha.Size = new System.Drawing.Size(102, 17);
            this.rbtResetSenha.TabIndex = 143;
            this.rbtResetSenha.TabStop = true;
            this.rbtResetSenha.Text = "Reset de Senha";
            this.rbtResetSenha.UseVisualStyleBackColor = true;
            this.rbtResetSenha.CheckedChanged += new System.EventHandler(this.rbtResetSenha_CheckedChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(125, 96);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(401, 36);
            this.label9.TabIndex = 145;
            this.label9.Text = "GERENCIAMENTO DE USUÁRIOS";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::CRMagazine.Properties.Resources.B1;
            this.pictureBox1.Location = new System.Drawing.Point(1, -2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(215, 95);
            this.pictureBox1.TabIndex = 144;
            this.pictureBox1.TabStop = false;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(204, 393);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(69, 20);
            this.label10.TabIndex = 147;
            this.label10.Text = "Estoque";
            // 
            // cbxEstoque
            // 
            this.cbxEstoque.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxEstoque.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxEstoque.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxEstoque.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxEstoque.FormattingEnabled = true;
            this.cbxEstoque.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxEstoque.Location = new System.Drawing.Point(343, 393);
            this.cbxEstoque.Name = "cbxEstoque";
            this.cbxEstoque.Size = new System.Drawing.Size(63, 24);
            this.cbxEstoque.TabIndex = 146;
            this.cbxEstoque.Text = "...";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(426, 311);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(52, 20);
            this.label11.TabIndex = 149;
            this.label11.Text = "Assist";
            this.label11.Visible = false;
            // 
            // cbxAssist
            // 
            this.cbxAssist.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxAssist.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxAssist.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxAssist.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxAssist.FormattingEnabled = true;
            this.cbxAssist.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxAssist.Location = new System.Drawing.Point(565, 311);
            this.cbxAssist.Name = "cbxAssist";
            this.cbxAssist.Size = new System.Drawing.Size(63, 24);
            this.cbxAssist.TabIndex = 148;
            this.cbxAssist.Text = "...";
            this.cbxAssist.Visible = false;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(204, 454);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(45, 20);
            this.label12.TabIndex = 151;
            this.label12.Text = "ADM";
            // 
            // cbxADM
            // 
            this.cbxADM.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbxADM.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbxADM.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cbxADM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cbxADM.FormattingEnabled = true;
            this.cbxADM.Items.AddRange(new object[] {
            "yes",
            "no"});
            this.cbxADM.Location = new System.Drawing.Point(343, 454);
            this.cbxADM.Name = "cbxADM";
            this.cbxADM.Size = new System.Drawing.Size(63, 24);
            this.cbxADM.TabIndex = 150;
            this.cbxADM.Text = "...";
            // 
            // frmInserirUsuario
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(624, 508);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.cbxADM);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.cbxAssist);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.cbxEstoque);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.rbtResetSenha);
            this.Controls.Add(this.btnLimparTela);
            this.Controls.Add(this.txtUsuario);
            this.Controls.Add(this.btnConcluir);
            this.Controls.Add(this.rbtDeleteUsuario);
            this.Controls.Add(this.rbtUsuarioExistente);
            this.Controls.Add(this.rbtNovoUsuario);
            this.Controls.Add(this.cbxUsuario);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbxAjuste);
            this.Controls.Add(this.cbxEmbalagem);
            this.Controls.Add(this.cbxConsultas);
            this.Controls.Add(this.cbxExpedicao);
            this.Controls.Add(this.cbxRunin);
            this.Controls.Add(this.cbxReparo);
            this.Controls.Add(this.cbxVistoria);
            this.Controls.Add(this.cbxEntrada);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmInserirUsuario";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gerenciar Usuarios";
            this.Load += new System.EventHandler(this.frmInserirUsuario_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbxEntrada;
        private System.Windows.Forms.ComboBox cbxVistoria;
        private System.Windows.Forms.ComboBox cbxReparo;
        private System.Windows.Forms.ComboBox cbxConsultas;
        private System.Windows.Forms.ComboBox cbxExpedicao;
        private System.Windows.Forms.ComboBox cbxRunin;
        private System.Windows.Forms.ComboBox cbxAjuste;
        private System.Windows.Forms.ComboBox cbxEmbalagem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbxUsuario;
        private System.Windows.Forms.RadioButton rbtNovoUsuario;
        private System.Windows.Forms.RadioButton rbtUsuarioExistente;
        private System.Windows.Forms.RadioButton rbtDeleteUsuario;
        private System.Windows.Forms.Button btnConcluir;
        private System.Windows.Forms.TextBox txtUsuario;
        private System.Windows.Forms.Button btnLimparTela;
        private System.Windows.Forms.RadioButton rbtResetSenha;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox cbxEstoque;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cbxAssist;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox cbxADM;

    }
}