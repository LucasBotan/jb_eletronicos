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
    public partial class frmProcessosDevolucao : Form
    {
        public frmProcessosDevolucao(string texto)
        {
            InitializeComponent();
            lblUsuario.Text = texto;
        }

        public string DevOuSuc = "";
        private void frmAjusteDevolucao_Load(object sender, EventArgs e)
        {
            txtOS.Select();
            rbtDevolucao.Checked = true;
            DevOuSuc = "DEVOLUÇÃO";
        }

        Conexao cx = new Conexao();
        Consulta consulta = new Consulta();
        Impressao imprimir = new Impressao();

        private void btnBuscarChamado_Click(object sender, EventArgs e)
        {
            if (txtOS.Text.Length > 0)
            {
                consulta.Coluna = "OS";
                consulta.ValorDesejado = txtOS.Text;
                consulta.ComData = "SIM";
                consulta.ConsultaTudo();
                if (consulta.Retorno == "ok")
                {
                    txtModelo.Text = consulta.Descricao;
                    txtStatus.Text = consulta.Status;
                    txtVarejista.Text = consulta.Varejista;
                    txtCodVarejo.Text = consulta.CodVarejo;
                    txtNS.Text = consulta.NS;

                    consulta.ConsultarCodVarejo(txtCodVarejo.Text, txtVarejista.Text);
                    if (consulta.Retorno == "ok")
                    {
                        txtEAN.Text = consulta.EANpuri;
                    }
                    //txtMotivoDevolucao.Select();
                    txtMotivoDevolucao.Text = DevOuSuc;
                    btnConcluir.Select();
                }
                else
                {
                    MessageBox.Show("OS NÃO ENCONTRADA");
                    txtOS.Text = "";
                    txtOS.Select();
                }
            }
            else
            {
                MessageBox.Show("INFORME UMA OS.");
            }
        }

       
        private void btnConcluir_Click(object sender, EventArgs e)
        {
            
            if (MessageBox.Show("DESEJA CONCLUIR A " + DevOuSuc + " ?", "PERGUNTA", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (txtOS.Text.Length == 0)
                {
                    MessageBox.Show("PEENCHA A OS.");
                }
                else if (txtModelo.Text.Length == 0)
                {
                    MessageBox.Show("INFORME O MODELO.");
                }
                else if (txtMotivoDevolucao.Text.Length == 0)
                {
                    MessageBox.Show("PEENCHA O MOTIVO DA " + DevOuSuc);
                }
                else if (chbNaoImprimir.Checked == false && rbt220.Checked == false && rbt110.Checked == false && rbtBIv.Checked == false)
                {
                    MessageBox.Show("SELECIONE A VOLTAGEM PARA IMPRESSÃO.");
                }
                else
                {
                    string Classificacao = "DEVOLUCAO";
                    if (rbtSucata.Checked)
                    {
                        Classificacao = "SUCATA";
                    }
                    else if (rbtDivergencia.Checked)
                    {
                        Classificacao = "DIVERGENCIA";
                    }
                    else if (rbtReprovado.Checked)
                    {
                        Classificacao = "REPROVADO";
                    }
                    consulta.comando = "";
                    consulta.comando += "update Chamados set status = 'EXPEDICAO', Classificacao = '" + Classificacao + "', MotivoDevolucao = '" + txtMotivoDevolucao.Text + "' where OS = '" + txtOS.Text + "'";
                    consulta.Atualizar();
                    if (consulta.Retorno == "ok")
                    {
                        //======Insere na tabela Historico==========================
                        string Local = Classificacao;
                        consulta.DataAtual();
                        consulta.InsereHistorico(txtOS.Text, lblUsuario.Text, Local, consulta.dataNormal, consulta.hora);
                        //=====fim da inserção======================================


                        if (txtVarejista.Text.Contains("MAGAZINE") 
                            || txtVarejista.Text.Contains("B2W") 
                            || txtVarejista.Text.Contains("SHOPLOKO") 
                            || txtVarejista.Text.Contains("LOJAS CEM")
                            || txtVarejista.Text.Contains("COMPROU 123")
                            )
                        {
                            ImprimirSaldoMagazine(Classificacao);
                        }

                        MessageBox.Show(DevOuSuc + " CONCLUIDA. \r\nENVIAR PARA EXPEDIÇÃO.");
                        consulta.LimparControles(this);
                        rbt110.Checked = false;
                        rbt220.Checked = false;
                        rbtBIv.Checked = false;
                        txtOS.Select();
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
            bool usarConfigDaImpressora = false;
            if (chbConfigImpressora.Checked)
            {
                usarConfigDaImpressora = true;
            }
            string varejista = "MAGAZINE LUIZA";
            if (!txtVarejista.Text.Contains("MAGAZINE"))
            {
                varejista = txtVarejista.Text;
            }
            imprimir.EtiquetaSaldoMagazine(usarConfigDaImpressora, Voltagem, txtCodVarejo.Text, consulta.Filial, txtEAN.Text, txtModelo.Text, txtNS.Text, Classificacao, varejista);
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

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("DESEJA LIMPAR A TELA?", "PERGUNTA", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                consulta.LimparControles(this);
                txtOS.Select();
            }
        }

        private void txtChamado_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnBuscarChamado.PerformClick();
            }
        }

        private void rbtSucata_CheckedChanged(object sender, EventArgs e)
        {
            if(rbtSucata.Checked)
            {
                DevOuSuc = "SUCATA";
                lblTitulo.Text = DevOuSuc;
            }
        }

        private void rbtDevolucao_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtDevolucao.Checked)
            {
                DevOuSuc = "DEVOLUÇÃO";
                lblTitulo.Text = DevOuSuc;
            }
        }

        private void rbtReprovado_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtReprovado.Checked)
            {
                DevOuSuc = "REPROVADO";
                lblTitulo.Text = DevOuSuc;
            }
        }

        private void rbtDivergencia_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtDivergencia.Checked)
            {
                DevOuSuc = "DIVERGENCIA";
                lblTitulo.Text = DevOuSuc;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //ImprimirSaldoMagazine("SUCATA");
        }



    }
}
