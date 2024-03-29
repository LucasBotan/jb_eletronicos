﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace CRMagazine
{
    public partial class frmEstoqueSolicitarPecasSAOPAULO : Form
    {
        public string tecnica = "";
        public frmEstoqueSolicitarPecasSAOPAULO(string texto, string OS, string Barebone, string Chamado, string Situacao)
        {
            InitializeComponent();
            lblUsuario.Text = texto;
            txtNS.Text = OS;
            txtChamado.Text = Chamado;
            txtBarebone.Text = Barebone;
            txtSituacao.Text = Situacao;
            if (OS != "")
            {
                btnAdicionar.Visible = true;
                btnRetirar.Visible = true;
                lstPedidos.Visible = true;
                btnConcluir.Enabled = false;
                btnLimparTela.Enabled = false;
                txtQntd.Text = "1";
                txtQntd.ReadOnly = true;
                tecnica = "yes";
                consultarEstoque();
            }
        }

        Conexao cx = new Conexao();
        Consulta consultar = new Consulta();

        private void frmEstoqueSolicitarPecas_Load(object sender, EventArgs e)
        {
            consultarEstoque();
            FormatarGrid();
            txtCodPeca.Select();
        }

        public void FormatarGrid()
        {
            DataGridViewImageColumn img = new DataGridViewImageColumn();
            //Image imagem = Image.FromFile(pictureBox1.Image);
            Image imagem = pictureBox2.Image;
            img.Image = imagem;
            dgvConsulta.Columns.Add(img);
            img.HeaderText = "Busca";
            img.Name = "Busca";

            //dgvConsulta.AutoResizeColumns();
            //dgvConsulta.Columns[0].Visible = false;
            dgvConsulta.RowHeadersVisible = false;
            // dgvConsulta.Columns[1].Width = 80;
            //dgvConsulta.Columns[2].Width = 80;
            //dgvConsulta.Columns[3].Width = 200;
            dgvConsulta.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgvConsulta.ScrollBars = ScrollBars.Vertical;
            //dgvConsulta.ScrollBars = System.Windows.Forms.ScrollBars.Both;
        }

        public void consultarEstoque()
        {
            string sql = "";
            if ((txtBarebone.Text.Length == 0) && (txtDescricao.Text.Length == 0))
            {
                sql = "";
                sql = "select Codigo, Descricao, Barebone, sum(Quantidade)As Quantidade from Estoque where Descricao like '%xxzzyy%' and Barebone like '%xxzzyy%' group by Codigo, Descricao, Barebone ";
                cx.ConectarSP();
                SqlDataAdapter da = new SqlDataAdapter(sql, cx.c);
                DataSet ds = new DataSet();
                da.Fill(ds, "Estoque");
                dgvConsulta.DataSource = ds.Tables["Estoque"];
                cx.Desconectar();
            }
            else
            {
                sql = "";
                sql = "select Codigo, Descricao, Barebone, sum(Quantidade)As Quantidade from Estoque where Descricao like '%" + txtDescricao.Text + "%' and Barebone like '%" + txtBarebone.Text + "%' group by Codigo, Descricao, Barebone ";
                cx.ConectarSP();
                SqlDataAdapter da = new SqlDataAdapter(sql, cx.c);
                DataSet ds = new DataSet();
                da.Fill(ds, "Estoque");
                dgvConsulta.DataSource = ds.Tables["Estoque"];
                cx.Desconectar();
            }            
        }


        public void ConsultaCodigo()
        {
            consultar.Codigo = txtCodPeca.Text.ToString();
            consultar.consultarEstoqueSP();
            if (consultar.Retorno == "ok")
            {
                txtDescPeca.Text = consultar.Descricao;
                txtTipo.Text = consultar.Tipo;
                if ((txtNS.Text == "") && (txtTipo.Text == "Bateria" || txtTipo.Text == "PLM"))
                {
                    MessageBox.Show("NÃO É POSSIVEL SOLICITAR PEÇAS FUNCIONAIS!");
                    txtCodPeca.Text = "";
                    txtDescricao.Text = "";
                    txtDescPeca.Text = "";
                    txtTipo.Text = "";
                    txtCodPeca.Select();
                }
                txtQntd.Select();
            }
            else
            {
                consultar.Codigo = txtCodPeca.Text;
                consultar.ConsultarCodigoSAP();
                if (consultar.Retorno == "ok")
                {
                    txtDescPeca.Text = consultar.Descricao;
                }
                else
                {
                    MessageBox.Show("CÓDIGO NÃO ENCONTRADO NO ESTOQUE NEM NO SAP!\r\n\r\nCHAMAR O SUPORTE.");
                    txtCodPeca.Text = "";
                    txtCodPeca.Select();
                }
            }

            //contar a qntd em estoque 
            int Estoque = 0;
            consultar.Codigo = txtCodPeca.Text.ToString();
            consultar.ContarEstoqueSP();
            Estoque = consultar.Contagem;
            //contar a qntd já solicitada
            int Pedidos = 0;
            consultar.Codigo = txtCodPeca.Text.ToString();
            consultar.ContarEmPedidosSP();
            Pedidos = consultar.Contagem;
            //contar a qntd pedida do magazine tmb
            int PedidosMagazine = 0;
            consultar.Codigo = txtCodPeca.Text.ToString();
            consultar.ContarEmPedidosMagazineSP();
            PedidosMagazine = consultar.Contagem;

            //subtrai as qntds para obter a disponibilidade real em estoque
            txtDisponivel.Text = Convert.ToString(Estoque - Pedidos - PedidosMagazine);
        }


        private void txtCodPeca_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
            {
                e.Handled = true;
            }
            if (e.KeyChar == 13)
            {
                ConsultaCodigo();
            }
        }

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            if ((Convert.ToInt32(txtDisponivel.Text) < Convert.ToInt32(txtQntd.Text)) || (Convert.ToInt32(txtQntd.Text) <= 0))
            {
                MessageBox.Show("Quantidade Indisponivel");
               // txtQntd.Text = "";
               // txtQntd.Select();
            }
            else if (txtCodAntigo.Text.Length == 0 || txtDescPecaAntiga.Text.Length == 0)
            {
                MessageBox.Show("PREENCHA O CÓDIGO ANTIGO E TECLE ENTER.\r\n\r\n(O MESMO DO ASSIST).");
            }
            else
            {
                AdicionarNaLista();
                btnConcluir.Enabled = true;
            }
           
        }

        public void AdicionarNaLista()
        {
            consultar.DataAtual();
            string data = consultar.dataCompleta;
            string Aux = "";
            //Aux += " '" + lblUsuario.Text + "', ";
            Aux += " '" + txtCodPeca.Text + "', ";
            Aux += " '" + txtDescPeca.Text + "', ";
            Aux += " '" + txtQntd.Text + "', ";
            Aux += " '" + txtCodAntigo.Text + "', ";
            Aux += " '" + txtDescPecaAntiga.Text + "', ";
            // Aux += "'AGUARDANDO', ";
            // Aux += " '" + data + "', ";
            // Aux += " '" + txtOS.Text + "' )";
            if (lstPedidos.Items.Contains(Aux))
            {
                MessageBox.Show("PEÇA JÁ INSERIDA NO PEDIDO");
            }
            else
            {
                lstPedidos.Items.Add(Aux);   
            }                    
        }

        private void btnRetirar_Click(object sender, EventArgs e)
        {
            RetirarDaLista();
        }

        public void RetirarDaLista()
        {
            //lstPedidos.SelectedItems.Clear();
            try
            {
                lstPedidos.Items.RemoveAt(lstPedidos.SelectedIndex);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);
            }
            
        }

        public void InserirAguardoBackUP()
        {
            consultar.DataAtual();
            string data = consultar.dataCompleta;
            string stringMontada = "";
            stringMontada = "Insert into AGUARDOBACKUP (Chamado, NS, Codigo, Descricao, Quantidade, CodAntigo, DescPecaAntiga, Status, DataPedido) Values ";
            int rows = lstPedidos.Items.Count;
            foreach (string item in lstPedidos.Items)
            {
                if (rows > 1)
                {
                    stringMontada += " ('" + txtChamado.Text + "',";
                    stringMontada += " '" + txtNS.Text + "',";
                    stringMontada += item;
                    stringMontada += "'PENDENTE', ";
                    stringMontada += " '" + data + "' ), ";
                }
                else
                {
                    stringMontada += " ('" + txtChamado.Text + "', ";
                    stringMontada += " '" + txtNS.Text + "',";
                    stringMontada += item;
                    stringMontada += "'PENDENTE', ";
                    stringMontada += " '" + data + "' ) ";
                }
                rows--;
            }
            consultar.comando = "";
            consultar.comando = stringMontada;
            consultar.Atualizar();
        }

        public void concluirPedido()
        {
            consultar.DataAtual();
            string data = consultar.dataCompleta;
            string stringMontada = "";
            stringMontada = "Insert into MovimentacaoEstoque (Usuario, Codigo, Descricao, Quantidade, CodAntigo, DescPecaAntiga, Status, DataPedido) Values ";
            int rows = lstPedidos.Items.Count;
            foreach (string item in lstPedidos.Items)
            {
                if (rows > 1)
                {
                    stringMontada += " ('MAGAZINE', ";
                    stringMontada += item;
                    stringMontada += "'AGUARDANDO', ";
                    stringMontada += " '" + data + "' ), ";                   
                }
                else
                {
                    stringMontada += " ('MAGAZINE', ";
                    stringMontada += item;
                    stringMontada += "'AGUARDANDO', ";
                    stringMontada += " '" + data + "' ) ";
                }
                rows--;
            }
            consultar.comando = "";
            consultar.comando = stringMontada;
            consultar.AtualizarSP();  
        }

        private void btnConcluir_Click(object sender, EventArgs e)
        {
            if ((Convert.ToInt32(txtDisponivel.Text) < Convert.ToInt32(txtQntd.Text)) || (Convert.ToInt32(txtQntd.Text) <= 0))
            {
                MessageBox.Show("Quantidade Indisponivel");
                txtQntd.Text = "";
                txtQntd.Select();
            }
            else if (lstPedidos.Items.Count == 0)
            {
                MessageBox.Show("INFORME A PEÇA ANTIGA.");
            }
            else
            {
                concluirPedido();
                InserirAguardoBackUP();
                if (consultar.Retorno == "ok")
                {
                    //======Insere na tabela Historico==========================
                    string Local = "AGUARDOBACKUP";
                    consultar.DataAtual();
                    consultar.InsereHistorico(txtNS.Text, lblUsuario.Text, Local, consultar.dataNormal, consultar.hora);
                    //=====fim da inserção======================================

                    consultar.comando = "update Chamados set Status = 'AGUARDOBACKUP' where NS = '" + txtNS.Text + "' AND STATUS = 'REPARO'";
                    consultar.Atualizar();
                    if (consultar.Retorno == "ok")
                    {
                        MessageBox.Show("EQUIPAMENTO ENVIADO PARA PP.");
                        Reset();
                    }
                   
                        this.Close();
                                     
                }
                else
                {
                    MessageBox.Show("ERRO: PEDIDO NÃO CONCLUIDO");
                }
            }
            
        }

        private void btnLimparTela_Click(object sender, EventArgs e)
        {
            Reset();
        }

        public void Reset()
        {
            consultar.LimparControles(this);
            txtCodPeca.Select();
        }

        private void txtQntd_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != (char)8)
            {
                e.Handled = true;
            }
        }

        private void btnConsultar_Click(object sender, EventArgs e)
        {
            consultarEstoque();
        }

        private void dgvConsulta_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dgvConsulta.Columns[e.ColumnIndex].Name == "Busca")
                {
                    txtCodPeca.Text = dgvConsulta.Rows[e.RowIndex].Cells["Codigo"].Value.ToString();
                    ConsultaCodigo();   
                   // txtDescricao.Text = dgvConsulta.Rows[e.RowIndex].Cells["Descricao"].Value.ToString();
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message);
            }
        }

        private void txtOS_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtCodAntigo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (txtCodAntigo.Text == "S/N")
                {
                    txtDescPecaAntiga.Text = "S/N";
                }
                else if (txtCodAntigo.Text.Length == 0)
                {
                    MessageBox.Show("INFORME O CÓDIGO DA PEÇA ANTIGA.");
                }
                else
                {
                    consultar.Codigo = txtCodAntigo.Text.ToString();
                    consultar.consultarEstoqueSP();
                    if (consultar.Retorno == "ok")
                    {
                        txtDescPecaAntiga.Text = consultar.Descricao;
                    }
                    else
                    {
                        consultar.Codigo = txtCodPeca.Text;
                        consultar.ConsultarCodigoSAP();
                        if (consultar.Retorno == "ok")
                        {
                            txtDescPecaAntiga.Text = consultar.Descricao;
                        }
                        else
                        {
                            MessageBox.Show("CÓDIGO NÃO ENCONTRADO NO ESTOQUE NEM NO SAP!\r\n\r\nCHAMAR O SUPORTE.");
                            txtCodPeca.Text = "";
                            txtCodPeca.Select();
                        }                        
                    }
                }
            }
        }
           
        

       


    }
}
