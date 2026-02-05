using System;
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
    public partial class frmExpedicaoExpedicao : Form
    {
        public frmExpedicaoExpedicao(string texto)
        {
            InitializeComponent();
            lblUsuario.Text = texto;
        }

        Conexao cx = new Conexao();
        Conexao consulte = new Conexao();
        Consulta consulta = new Consulta();

        private void frmExpedicaoExpedicao_Load(object sender, EventArgs e)
        {
            txtOS.Select();
          //  ListarTudo();
          //  FormatarGrid();
            ContadorDeProducao();
        }

        public void ContadorDeProducao()
        {
            consulta.DataAtual();
            string Status = "EXPEDIDO";
            consulta.comando = "select COUNT(*) as QNT from Historico where Usuario = '" + lblUsuario.Text + "' and Status = '" + Status + "' and Data = '" + consulta.dataNormal + "'";
            consulta.consultarHistorico();
            lblContador.Text = consulta.cont.ToString();
        }

        public void FormatarGrid()
        {
            dgvParaAbrir.Columns[0].Visible = false;
            dgvParaAbrir.RowHeadersVisible = false;
            // dgvConsulta.Columns[1].Width = 60;
            //dgvParaAbrir.Columns[1].Width = 200;
            dgvParaAbrir.Columns["Descricao"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //dgvConsulta.ScrollBars = ScrollBars.Vertical;
            dgvParaAbrir.ScrollBars = System.Windows.Forms.ScrollBars.Both;
        }

        public void ListarTudo()
        {
            string sql = "";
            sql += " Select idChamados, OS, NS, Descricao, Status, DataEntrada, DATEDIFF(DAY, DataEntrada, GETDATE()) as DNP, StatusA1 From Chamados where Status = 'EXPEDICAO'";
            cx.Conectar();
            SqlDataAdapter da = new SqlDataAdapter(sql, cx.c);
            DataSet ds = new DataSet();
            da.Fill(ds, "Chamados");
            dgvParaAbrir.DataSource = ds.Tables["Chamados"];
            cx.Desconectar();
            lblTotal.Text = dgvParaAbrir.Rows.Count.ToString();
            //listar datas     
        }

        private void btnConcluir_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            for (int i = 0; i < dgvParaAbrir.RowCount; i++)
            {
                string OS = dgvParaAbrir.Rows[i].Cells["OS"].Value.ToString();


                // ========================= ATUALIZA BANCO CHAMADOS ============================
                consulta.comando = "update Chamados Set Status = 'FINALIZADO', DataSaida = '" + consulta.data + "' where Status = 'EXPEDICAO' and OS = '" + OS + "'";
                consulta.Atualizar();
                // ==============================================================================

                //======Insere na tabela Historico==========================
                string StatusHistorico = "EXPEDIDO";
                consulta.DataAtual();
                consulta.InsereHistorico(OS, lblUsuario.Text, StatusHistorico, consulta.dataNormal, consulta.hora);
                //=====fim da inserção======================================

            }
            

            MessageBox.Show("EQUIPAMENTOS EPEDIDOS.");
            this.Cursor = Cursors.Default;
            ListarTudo();
            ContadorDeProducao();
        }

        private void btnBusca_Click(object sender, EventArgs e)
        {
            consulta.Coluna = "OS";
            consulta.ValorDesejado = txtOS.Text;
          //  consulta.ComData = "SIM";  // para não puxar os finalizados
            // consulta.DataDesejada = txtDataDesejadaNaoLimpar.Text;
            consulta.ConsultaTudo();
            if (consulta.Retorno == "falha")
            {
                consulta.PlayFail();
                MessageBox.Show("OS NÃO ENCONTRADA.");
                txtOS.Select();
                txtOS.SelectAll();
            }
            else if (consulta.Status == "FINALIZADO")
            {
                consulta.PlayFail();
                MessageBox.Show("OS JÁ EXPEDIDA.");
                txtOS.Select();
                txtOS.SelectAll();
            }
            else if (consulta.Status != "EXPEDICAO")
            {
                consulta.PlayFail();
                MessageBox.Show("OS NÃO ESTA EM EXPEDICAO.\r\nSTATUS = " + consulta.Status);
                txtOS.Select();
                txtOS.SelectAll();
            }
            else if ((consulta.Varejista != "MULTIVAREJO" && consulta.Classificacao == "ORCAMENTO" && !consulta.StatusA1.Contains("APROVADO")) && chbSemA1.Checked == false)
            {
                consulta.PlayFail();
                MessageBox.Show("OS NÃO APROVADA.");
                txtOS.Select();
                txtOS.SelectAll();
            }
            else
            {

               
                txtDescricao.Text = consulta.Descricao;
                txtClassificacao.Text = consulta.Classificacao;
                txtStatusA1.Text = consulta.StatusA1;
                txtCodVarejo.Text = consulta.CodVarejo;
                txtVarejista.Text = consulta.Varejista;

                if (txtVarejista.Text == "MULTIVAREJO" || txtVarejista.Text == "LOJAS CEM")
                {
                    txtNF.Text = consulta.NotaFiscal;
                }
                else
                {
                    BuscarNF();
                }
                if (txtNF.Text != "0" && txtNF.Text.Length > 0)
                {
                    if (MessageBox.Show("CONCLUIR? ", "Pergunta", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        BaixarQntNF();
                        // ========================= ATUALIZA BANCO CHAMADOS ============================

                        string OS = consulta.Prevent(txtOS.Text);
                        string pendente = "PENDENTE";
                        if (txtNF.Text == "0")
                        {
                            NF = "SEM NOTA";
                            pendente = "SEM NOTA";
                        }
                        //consulta.comando = "update Chamados Set Status = 'FINALIZADO', Classificacao = 'DEVOLUCAO', MotivoDevolucao = 'DEVOLUÇÃO EXIGIDA', DataSaida = '" + consulta.data + "', NotaFiscal = '" + NF + "', NotaFiscalSaida = '" + pendente + "' where Status = 'EXPEDICAO' and OS = '" + txtOS.Text + "'";
                        consulta.comando = $"update Chamados Set Status = 'FINALIZADO', DataSaida = '{consulta.data}', NotaFiscal = '{NF}', NotaFiscalSaida = '{pendente}' where Status = 'EXPEDICAO' and OS = '{OS}'";
                        consulta.Atualizar();
                        // ==============================================================================
                        if (consulta.Retorno == "ok")
                        {
                            //======Insere na tabela Historico==========================
                            string StatusHistorico = "EXPEDIDO";
                            consulta.DataAtual();
                            consulta.InsereHistorico(OS, lblUsuario.Text, StatusHistorico, consulta.dataNormal, consulta.hora);
                            //=====fim da inserção======================================

                            lstContagem.Items.Add(OS);
                            int rows = lstContagem.Items.Count;
                            lblContagem.Text = rows.ToString();

                            ContadorDeProducao();
                            consulta.PlayOK();
                            LimparTela();
                        }

                    }
                    else
                    {
                        LimparTela();
                    }
                }
                else
                {
                    consulta.PlayFail();
                    MessageBox.Show("NÃO FOI POSSÍVEL ATRIBUIR UMA NOTA FISCAL.");
                    LimparTela();
                }

                
            }

        }

        public void LimparTela()
        {
            NF = "";
            consulta.LimparControles(this);
            txtOS.Select();
        }

        private void txtConcluir_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            lstContagem.Items.Clear();
            int rows = lstContagem.Items.Count;
            lblContagem.Text = rows.ToString();
            LimparTela();
        }

        private void txtOS_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnBusca.PerformClick();
            }
        }

        // PUXAR A NOTA FISCAL MAIS ANTIGA QUE CONTENHA O CODVAREJO E PRECISE SER BAIXADA A QUANTIDADE
        public void BuscarNF()
        {
            string codigo = consulta.Prevent(txtCodVarejo.Text);
            string varejista = consulta.Prevent(txtVarejista.Text);

            consulta.comando = $"select top(1) notafiscal as Quantidade from NotaFiscal where CodVarejo = '{codigo}' and QntRestante > 0 and Varejista = '{varejista}' order by CONVERT(date, Data, 103) ASC";
            consulta.consultarSimNao();
            txtNF.Text = consulta.qntNaPosicao;
        }

        public string NF = ""; 
        // BAIXAR QNT DA NOTA FISCAL MAIS ANTIGA QUE CONTENHA O CODVAREJO E PRECISE SER BAIXADA A QUANTIDADE
        public void BaixarQntNF()
        {
            string codigo = consulta.Prevent(txtCodVarejo.Text);
            string varejista = consulta.Prevent(txtVarejista.Text);
            string nota = consulta.Prevent(txtNF.Text);

            if (txtVarejista.Text == "MULTIVAREJO" || txtVarejista.Text == "LOJAS CEM")
            {
                NF = nota;
            }
            else
            {
                consulta.comando = $"select top(1) notafiscal as Quantidade from NotaFiscal where CodVarejo = '{codigo}' and QntRestante > 0 and Varejista = '{varejista}' order by CONVERT(date, Data, 103) ASC";
                consulta.consultarSimNao();
                NF = consulta.qntNaPosicao;
            }
            

            consulta.comando = $"update notafiscal set QntRestante = QntRestante - 1 where notafiscal = '{NF}' and CodVarejo = '{codigo}' and QntRestante > 0 and Varejista = '{varejista}'";
            consulta.Atualizar();
        }



    }
}
