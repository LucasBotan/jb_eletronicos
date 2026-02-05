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
using System.Drawing.Printing;

namespace CRMagazine
{
    public partial class frmAjustesAlterarClassificacao : Form
    {
        public frmAjustesAlterarClassificacao(string user, string OS)
        {
            InitializeComponent();
            lblUsuario.Text = user;
            txtOS.Text = OS;
        }

        Conexao cx = new Conexao();
        Consulta consulta = new Consulta();
        Impressao imprimir = new Impressao();
        private void frmAjustesAlterarClassificacao_Load(object sender, EventArgs e)
        {

        }

        private void btnLimparTela_Click(object sender, EventArgs e)
        {
            consulta.LimparControles(this);
            cbxClassificacao.Text = null;
            if (txtOS.Text.Length > 0)
            {
                txtOS.Focus();
                SendKeys.Send("{ENTER}");
            }
            else
            {
                txtOS.Focus();
            }
            
        }

        private void btnConcluir_Click(object sender, EventArgs e)
        {
            if (cbxClassificacao.Text == txtClassificacaoAtual.Text)
            {
                consulta.PlayFail();
                MessageBox.Show("A CLASSIFICAÇÃO NÃO PODE SER A MESMA.");
            }
            else if (txtSKU.Text.Length == 0)
            {
                consulta.PlayFail();
                MessageBox.Show("INFORME O CHAMADO.");
            }
            else if (cbxClassificacao.Text.Length == 0)
            {
                consulta.PlayFail();
                MessageBox.Show("INFORME A CLASSIFICAÇÃO.");
            }
            else if (chbNaoImprimir.Checked == false && rbt220.Checked == false && rbt110.Checked == false && rbtBIv.Checked == false)
            {
                MessageBox.Show("SELECIONE A VOLTAGEM PARA IMPRESSÃO.");
            }
            /*
            else if (txtMotivoAlteracaoClass.Text.Length == 0)
            {
                consulta.PlayFail();
                MessageBox.Show("INFORME O MOTIVO PARA A ALTERAÇÃO DA CLASSIFICAÇÃO.");
            }
            */
            else
            {
                if (MessageBox.Show("DESEJA CONCLUIR A ALTERAÇÃO DA CLASSIFICAÇÃO? \n ", "PERGUNTA", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    DateTime agora = DateTime.Now;
                    string data = agora.ToString();

                    consulta.comando = "update Chamados set Classificacao = '" + cbxClassificacao.Text + "' where OS = '" + txtOS.Text + "' and Status != 'FINALIZADO'";
                    consulta.Atualizar();
                    // MessageBox.Show(consulta.comando);                           
                    if (consulta.LinhasAfetadas > 0)
                    {
                        //======Insere na tabela Historico==========================
                        string StatusHistorico = "ALTERADOCLASSIFICACAO";
                        consulta.DataAtual();
                       //consulta.ConsultaPontuacao(Tipo, StatusHistorico);
                        consulta.InsereHistorico(txtOS.Text, lblUsuario.Text, StatusHistorico, consulta.dataNormal, consulta.hora);
                        //=====fim da inserção======================================

                        if ((
                            txtVarejista.Text.Contains("MAGAZINE") 
                            || txtVarejista.Text.Contains("B2W") 
                            || txtVarejista.Text.Contains("SHOPLOKO") 
                            || txtVarejista.Text.Contains("LOJAS CEM")
                            || txtVarejista.Text.Contains("COMPROU 123")
                            ) && chbNaoImprimir.Checked == false)
                        {
                            ImprimirSaldoMagazine(cbxClassificacao.Text);
                        }

                        consulta.LimparControles(this);
                        cbxClassificacao.Text = null;
                        txtOS.Focus();
                        consulta.PlayOK();
                        MessageBox.Show("ALTERAÇÃO CONCLUIDA!");
                    }
                    else
                    {
                        MessageBox.Show("Não foi possivel Alterar, Chamar o Suporte para Analise.");
                    }
                }
            }                  
        }

        public void ImprimirSaldoMagazine(string Classificacao)
        {
            string Voltagem = "";
            if (rbt110.Checked)
            {
                Voltagem = "110";
            }
            else if (rbtBIv.Checked)
            {
                Voltagem = "BI";
            }
            else
            {
                Voltagem = "220";
            }

            string varejista = "MAGAZINE LUIZA";
            if (!txtVarejista.Text.Contains("MAGAZINE"))
            {
                varejista = txtVarejista.Text;
            }

            bool usarConfigDaImpressora = false;
            if (chbConfigImpressora.Checked)
            {
                usarConfigDaImpressora = true;
            }
            imprimir.EtiquetaSaldoMagazine(usarConfigDaImpressora, Voltagem, txtCodVarejo.Text, consulta.Filial, txtEAN.Text, txtDescricao.Text, txtNS.Text, Classificacao, varejista);
            string codZPL = imprimir.s;
            // SELECIONAR IMPRESSORA OU UTILIZAR A PADRÃO
            if (chbSelecionarImpressora.Checked)
            {
                // Allow the user to select a printer.
                PrintDialog pd = new PrintDialog();
                pd.PrinterSettings = new PrinterSettings();
                if (DialogResult.OK == pd.ShowDialog(this))
                {
                    // Send a printer-specific to the printer.
                    RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, codZPL);
                }
            }
            else
            {
                string nomeImpressoraPadrao = (new PrinterSettings()).PrinterName;
                RawPrinterHelper.SendStringToPrinter(nomeImpressoraPadrao, codZPL);
            }
        }

        private void btnBuscarChamado_Click(object sender, EventArgs e)
        {
            consulta.Coluna = "OS";
            consulta.ValorDesejado = txtOS.Text;
            consulta.ComData = "SIM"; // para não puxar os finalizados
            consulta.ConsultaTudo();

            if (consulta.Retorno == "ok")
            {
                txtSKU.Text = consulta.SKU;
                txtDescricao.Text = consulta.Descricao;
                txtStatus.Text = consulta.Status;
                txtClassificacaoAtual.Text = consulta.Classificacao;
                Tipo = consulta.TipoEquip;
                txtVarejista.Text = consulta.Varejista;
                txtCodVarejo.Text = consulta.CodVarejo;
                txtNS.Text = consulta.NS;

                consulta.ConsultarCodVarejo(txtCodVarejo.Text, txtVarejista.Text);
                if (consulta.Retorno == "ok")
                {
                    txtEAN.Text = consulta.EANpuri;
                }

                cbxClassificacao.Select();
            }
            else
            {
                MessageBox.Show("NS NÃO ENCONTRADA.");
                consulta.LimparControles(this);
                txtOS.Text = "";
                Tipo = "";
                txtDescricao.Text = "";
                txtClassificacaoAtual.Text = "";
                txtSKU.Text = "";
                txtOS.Focus();
            }
        }

        public string Tipo = "";
        private void txtOS_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnBuscarChamado.PerformClick();
            }
        }


        private void cbxClassificacao_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxClassificacao.Text == txtClassificacaoAtual.Text && cbxClassificacao.Text.Length > 0)
            {
                consulta.PlayFail();
                MessageBox.Show("A CLASSIFICAÇÃO NÃO PODE SER A MESMA.");
                cbxClassificacao.Text = null;
            }
        }

        private void cbxClassificacao_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                //btnConcluir.Select();
                txtMotivoAlteracaoClass.Select();
            }
        }

        private void txtMotivoAlteracaoClass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnConcluir.Select();                
            }
        }

        private void rbt220_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chbNaoImprimir_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtOS_TextChanged(object sender, EventArgs e)
        {

        }

        
    }
}
