using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRMagazine
{
    public partial class frmHigienizacaoProdutoMontado : Form
    {
        public frmHigienizacaoProdutoMontado()
        {
            InitializeComponent();
        }

        Conexao cx = new Conexao();
        Consulta consulta = new Consulta();
        Impressao imprimir = new Impressao();

        private void frmHigienizacaoProdutoMontado_Load(object sender, EventArgs e)
        {
            txtOS.Select();
            //rbt200dpi.Checked = true;
            rbt220.Checked = true;
            PreencherComboboxStatus();
        }

        public void PreencherComboboxStatus()
        {
            SqlDataAdapter da;
            DataSet ds = new DataSet();
            string sql = "";
            sql = "";
            sql += " SELECT DISTINCT Usuario FROM Usuarios WHERE Embalagem = 'yes'";
            cx.Conectar();
            da = new SqlDataAdapter(sql, cx.c);
            da.Fill(ds, "Usuarios");
            cboUsuario.ValueMember = "idUsuario";
            cboUsuario.DisplayMember = "Usuario";
            cboUsuario.DataSource = ds.Tables["Usuarios"];
            cx.Desconectar();
            cboUsuario.Text = null;
            cboUsuario.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboUsuario.AutoCompleteSource = AutoCompleteSource.ListItems;
        }

        public void ContadorDeProducao()
        {
            if (cboUsuario.Text.Length > 0)
            {
                consulta.DataAtual();
                consulta.comando = "select COUNT(*) as Quantidade from Producao where Higienizador = '" + cboUsuario.Text + "' and Status = 'RECEBIDO' AND Data = '" + consulta.dataNormal + "'";
                consulta.consultarHistorico();
                lblContador.Text = consulta.cont.ToString(); 
            }
        }

        private void cboUsuario_SelectedValueChanged(object sender, EventArgs e)
        {
            lblUsuario.Text = cboUsuario.Text;
            ContadorDeProducao();
        }

        private void btnBusca_Click(object sender, EventArgs e)
        {
            consultarProduto();
        }

        private void txtOS_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 13) return;

            consultarProduto();
        }

        public void consultarProduto()
        {
            if (string.IsNullOrWhiteSpace(txtOS.Text))
            {
                MessageBox.Show($"INFORME A OS.");
                return;
            }

            bool existe = consulta.ConsultaProdutoMontado("OS", txtOS.Text, true);
            if (!existe)
            {
                MessageBox.Show("EQUIPAMENTO NÃO ENCONTRADO.");
                return;
            }

            if(consulta.status_produtoMontado != "HIGIENIZACAO")
            {
                MessageBox.Show($"EQUIPAMENTO NÃO ESTA EM HIGIENIZAÇÃO.\r\nSTATUS ATUAL = {consulta.status_produtoMontado}");
                return;
            }

            txtDescricao.Text = consulta.Descricao_produtoMontado;
            txtEAN.Text = consulta.EAN_produtoMontado;
            txtCodPositivo.Text = consulta.SKU_produtoMontado;
            txtTecnico.Text = consulta.tecnico_produtoMontado;
            txtCodigoVarejista.Text = consulta.codVarejista_produtoMontado;

        }

        private void btnConcluir_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cboUsuario.Text))
            {
                MessageBox.Show($"INFORME O USUÁRIO QUE FARÁ A HIGIENIZAÇÃO.");
                return;
            }
            if (string.IsNullOrWhiteSpace(txtOS.Text))
            {
                MessageBox.Show($"INFORME A OS.");
                return;
            }
            if (string.IsNullOrWhiteSpace(txtDescricao.Text))
            {
                MessageBox.Show($"INFORME O EQUIPAMENTO.");
                return;
            }

            if (string.IsNullOrWhiteSpace(cboClassificacao.Text))
            {
                MessageBox.Show($"INFORME A CLASSIFICAÇÃO.");
                return;
            }
            consulta.DataAtual();

            string status = "AGUARDANDO";
            consulta.comando = $"update Producao set Status = '{status}'" +
                $", Classificacao = '{cboClassificacao.Text}' " +
                $", Higienizador = '{cboUsuario.Text}' " +
                $", DataHigienizacao = '{consulta.dataHora}' " +
                $"where OS = '{txtOS.Text}' AND STATUS = 'HIGIENIZACAO'";
            consulta.Atualizar();
            if(consulta.LinhasAfetadas == 0)
            {
                MessageBox.Show($"ERRO AO ATUALIZAR.");
                return;
            }

            if (chbNaoImprimir.Checked == false)
            {
                Imprimir();
            }

            btnLimpar.PerformClick();
            MessageBox.Show("HIGIENIZAÇÃO CADASTRADA COM SUCESSO.");
            txtOS.Select();
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            consulta.LimparControles(this);
        }

        public void Imprimir()
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


            if (chbConfigImpressora.Checked)
            {
                imprimir.EtiquetaEANPuriConfig(Voltagem, txtCodigoVarejista.Text, txtCodPositivo.Text, txtEAN.Text, txtDescricao.Text, cboClassificacao.Text);
            }
            else
            {
                imprimir.EtiquetaEANPuri(Voltagem, txtCodigoVarejista.Text, txtCodPositivo.Text, txtEAN.Text, txtDescricao.Text,"");
            }

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

        
    }
}
