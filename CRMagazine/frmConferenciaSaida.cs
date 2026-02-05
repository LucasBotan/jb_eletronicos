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
    public partial class frmConferenciaSaida : Form
    {
        public frmConferenciaSaida()
        {
            InitializeComponent();
        }

        Consulta consulta = new Consulta();
        Conexao cx = new Conexao();

        public string inicio = "";

        private void frmConferenciaSaida_Load(object sender, EventArgs e)
        {
            PreencherCboVarejista();
            inicio = "ok";
        }

        public void PreencherCboVarejista()
        {
            SqlDataAdapter da;
            DataSet ds = new DataSet();
            string sql = "";
            sql = "";
            sql += " Select DISTINCT Varejista from ConfereNFSaida where Conferir != 0";
            cx.Conectar();
            da = new SqlDataAdapter(sql, cx.c);
            da.Fill(ds, "ConfereNFSaida");
            cboVarejista.ValueMember = "idConfereNFSaida";
            cboVarejista.DisplayMember = "Varejista";
            cboVarejista.DataSource = ds.Tables["ConfereNFSaida"];
            cx.Desconectar();
            cboVarejista.Text = null;
            cboVarejista.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cboVarejista.AutoCompleteSource = AutoCompleteSource.ListItems;
            // cboNotaFiscal.Text = "";
            // lblRestanteNF.Text = "0";
        }

        public void PreencherCboNotaFiscal()
        {
            if (chbComDT.Checked && Data == "")
            {
                MessageBox.Show("PREENCHA A DATA DA NOTA FISCAL.");
                cboNotaFiscal.DataSource = null;
            }
            else
            {
                SqlDataAdapter da;
                DataSet ds = new DataSet();
                string sql = "";
                sql = "";
                sql += " Select DISTINCT NotaFiscal from ConfereNFSaida where Conferir != 0 and Varejista = '" + cboVarejista.Text + "'";
                if (chbComDT.Checked && Data != "")
                {
                    sql += " and Data = '" + Data + "'";
                }
                cx.Conectar();
                da = new SqlDataAdapter(sql, cx.c);
                da.Fill(ds, "ConfereNFSaida");
                cboNotaFiscal.ValueMember = "idConfereNFSaida";
                cboNotaFiscal.DisplayMember = "NotaFiscal";
                cboNotaFiscal.DataSource = ds.Tables["ConfereNFSaida"];
                cx.Desconectar();
                cboNotaFiscal.Text = null;
                cboNotaFiscal.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cboNotaFiscal.AutoCompleteSource = AutoCompleteSource.ListItems;
                // cboNotaFiscal.Text = "";
                lblPendenteNF.Text = "PENDENTES: 0";
            }
        }

        private void txtCodVarejo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnBusca.PerformClick();
            }
        }

        private void btnBusca_Click(object sender, EventArgs e)
        {
            if (txtCodVarejo.Text.Length == 0)
            {
                consulta.PlayFail();
                MessageBox.Show("PREENCHA O CÓDIGO.");
                txtCodVarejo.Select();
                txtCodVarejo.SelectAll();
                btnConcluir.Visible = false;
            }
            else if (notas == "")
            {
                consulta.PlayFail();
                MessageBox.Show("INFORME AS NOTAS.");
                txtCodVarejo.Select();
                txtCodVarejo.SelectAll();
                btnConcluir.Visible = false;
            }
            else
            {
                consulta.ConsultaNotaFiscalSaida(cboVarejista.Text, "CodVarejo", txtCodVarejo.Text, notas);
                if (consulta.Retorno == "ok")
                {
                    txtDescricao.Text = consulta.NF_Descricao_Saida;
                    txtCodigo.Text = consulta.NF_Codigo_Saida;
                    txtEanConsulta.Text = consulta.NF_EAN_Saida;
                    txtSomaDasQuantidades.Text = consulta.NF_Qnt_somada;
                    //btnConcluir.Select();
                    //AKI
                    consulta.comando = "";
                    consulta.comando = "";
                    consulta.comando = "SELECT count(NotaFiscal) as Quantidade FROM ConfereNFSaida where idConfereNFSaida = (select top 1 idConfereNFSaida from ConfereNFSaida where CodVarejo = '" + txtCodigo.Text + "' and Conferir > 0 and Varejista = '" + cboVarejista.Text + "')";
                    consulta.consultarSimNao();
                    if (Convert.ToInt32(consulta.qntNaPosicao) > 0)
                    {
                        btnConcluir.Visible = true;
                        btnConcluir.Select();
                    }
                    else
                    {
                        consulta.PlayFail();
                        MessageBox.Show("QUANTIDADE JÁ ZERADA PARA ESSE CÓDIGO.");
                        txtCodVarejo.Select();
                        txtCodVarejo.SelectAll();
                        btnConcluir.Visible = false;

                    }
                }
                else
                {
                    consulta.PlayFail();
                    MessageBox.Show("NOTA FISCAL SEM QUANTIDADE PARA ESSE CÓDIGO VAREJO.");
                    txtCodVarejo.Select();
                    txtCodVarejo.SelectAll();
                    btnConcluir.Visible = false;
                }
            }            
        }

        

        private void txtEAN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnBuscaEAN.PerformClick();
            }
        }

        private void btnBuscaEAN_Click(object sender, EventArgs e)
        {
            if (txtEAN.Text.Length == 0)
            {
                consulta.PlayFail();
                MessageBox.Show("PREENCHA A EAN.");
                txtEAN.Select();
                btnConcluir.Visible = false;
            }
            else if (notas == "")
            {
                consulta.PlayFail();
                MessageBox.Show("INFORME AS NOTAS.");
                txtCodVarejo.Select();
                txtCodVarejo.SelectAll();
                btnConcluir.Visible = false;
            }
            else
            {
                //consultar o código pelo EAN, e utilizar o código para buscar as informações (preencher o campo código com o retorno da busca)
                consulta.ConsultarEAN("EAN", txtEAN.Text, "NÃO");
                txtCodVarejo.Text = consulta.CodVarejo;
                consulta.ConsultaNotaFiscalSaida(cboVarejista.Text, "CodVarejo", txtCodVarejo.Text, notas);

                //consulta.ConsultaNotaFiscalSaida(cboVarejista.Text, "EAN", txtEAN.Text, notas); // ANTES - NÃO DA PRA USAR ASSIM PQ AS NFS DE SAIDA GERADAS PELO BARBOSA NÃO TEM EAN
                if (consulta.Retorno == "ok")
                {
                    txtDescricao.Text = consulta.NF_Descricao_Saida;
                    txtCodigo.Text = consulta.NF_Codigo_Saida;
                    txtEanConsulta.Text = consulta.NF_EAN_Saida;
                    txtSomaDasQuantidades.Text = consulta.NF_Qnt_somada;
                    //btnConcluir.Select();
                    //AKI
                    consulta.comando = "";
                    consulta.comando = "";
                    consulta.comando = "SELECT count(NotaFiscal) as Quantidade FROM ConfereNFSaida where idConfereNFSaida = (select top 1 idConfereNFSaida from ConfereNFSaida where CodVarejo = '" + txtCodigo.Text + "' and Conferir > 0 and Varejista = '" + cboVarejista.Text + "')";
                    consulta.consultarSimNao();
                    if (Convert.ToInt32(consulta.qntNaPosicao) > 0)
                    {
                        btnConcluir.Visible = true;
                        btnConcluir.Select();
                    }
                    else
                    {
                        consulta.PlayFail();
                        MessageBox.Show("QUANTIDADE JÁ ZERADA PARA ESSE CÓDIGO.");
                        txtEAN.SelectAll();
                        txtEAN.Select();
                        btnConcluir.Visible = false;
                    }
                }
                else
                {
                    consulta.PlayFail();
                    MessageBox.Show("NOTA FISCAL SEM QUANTIDADE PARA ESSE CÓDIGO VAREJO.");
                    btnLimpar.PerformClick();
                    txtEAN.SelectAll();
                    txtEAN.Select();
                    btnConcluir.Visible = false;
                }
            }
        }

        private void cboVarejista_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cboVarejista.Text.Length > 0 && inicio != "")
            {
                lstColunas.Items.Clear();
                ContarVarejista();
                PreencherCboNotaFiscal();
            }
        }

        private void chbComDT_CheckedChanged(object sender, EventArgs e)
        {
            if (chbComDT.Checked)
            {
                lstColunas.Items.Clear();
                ContarVarejista();
                PreencherCboNotaFiscal();
                mtbDataNF.Visible = true;
                Data = "";
                mtbDataNF.Select();
                lnkListarPorData.Visible = true;
            }
            else
            {
                lstColunas.Items.Clear();
                ContarVarejista();
                PreencherCboNotaFiscal();
                mtbDataNF.Visible = false;
                Data = "";
                mtbDataNF.Text = "";
                lnkListarPorData.Visible = false;
            }
        }

        public void ContarVarejista()
        {
            if (cboVarejista.Text.Length > 0)
            {
                //===============verificando restante ============
                string comando = "";
                comando += "SELECT sum(convert(numeric,Conferir)) as Quantidade from ConfereNFSaida where Varejista = '" + cboVarejista.Text + "' and Conferir > 0";
                if (chbComDT.Checked && Data != "")
                {
                    comando += " and Data = '" + Data + "'";
                }
                consulta.comando = comando;
                consulta.consultarSimNao();
                if (consulta.qntNaPosicao == "")
                {
                    consulta.qntNaPosicao = "0";
                }
                lblRestanteVarejista.Text = consulta.qntNaPosicao.ToString();

                //================================================                     
            }
        }

        private void cboNotaFiscal_SelectedValueChanged(object sender, EventArgs e)
        {

        }

        private void cboNotaFiscal_KeyDown(object sender, KeyEventArgs e)
        {
            MessageBox.Show("ESCRITA NÃO PERMITIDA.");
            if (cboNotaFiscal.Text.Length > 0)
            {
                cboNotaFiscal.Text = "";
            }
        }

        private void lnkAtualizar_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            inicio = "";
            PreencherCboVarejista();
            inicio = "ok";
            AtualizaContadores();
            limpar();
            btnLimparLista.PerformClick();
            chbComDT.Checked = false;
        }

        public void AtualizaContadores()
        {
            if (cboVarejista.Text.Length > 0)
            {
                consulta.comando = "SELECT sum(convert(numeric,Conferir)) as Quantidade from ConfereNFSaida where Conferir > 0 and Varejista = '" + cboVarejista.Text + "' ";
                consulta.consultarSimNao();
                if (consulta.qntNaPosicao == "")
                {
                    consulta.qntNaPosicao = "0";
                }
                lblRestanteVarejista.Text = consulta.qntNaPosicao.ToString();

                AtualizarListaNotas();
            }
            else
            {
                MessageBox.Show("SELECIONE UM VAREJISTA.");
                lblRestanteVarejista.Text = "0";
            }
        }

        public void AtualizarListaNotas()
        {
            if (lstColunas.Items.Count > 0)
            {
                consulta.comando = "SELECT sum(convert(numeric,Conferir)) as Quantidade from ConfereNFSaida where NotaFiscal in (" + notas + ") and Conferir > 0";
                consulta.consultarSimNao();
                if (consulta.qntNaPosicao == "")
                {
                    consulta.qntNaPosicao = "0";
                }
                lblPendenteNF.Text = "PENDENTES: " + consulta.qntNaPosicao.ToString();
            }
            else
            {
                lblPendenteNF.Text = "PENDENTES: 0";
            }
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            limpar();
        }

        public void limpar()
        {
            txtEAN.Text = "";
            txtCodVarejo.Text = "";
            txtDescricao.Text = "";
            txtCodigo.Text = "";
            txtEanConsulta.Text = "";
            txtQtd.Text = "";
            txtSomaDasQuantidades.Text = "";
            btnConcluir.Visible = false;
        }

        private void btnListarNF_Click(object sender, EventArgs e)
        {
            if (cboVarejista.Text.Length > 0 && lstColunas.Items.Count > 0)
            {
                ListarTudo(false);
            }
            else
            {
                MessageBox.Show("INFORME O VAREJISTA E A NOTA FISCAL.");
            }
        }

        public void ListarTudo(bool ComData)
        {
            string sql = "";
            if (ComData == false)
            {
                sql += " Select NotaFiscal as NF, CodVarejo as Cod, QntProdutos as QNT, Conferir as Rest, Descricao as Descr From ConfereNFSaida where Varejista = '" + cboVarejista.Text + "' and NotaFiscal in (" + notas + ") ";
            }
            else
            {
                sql += " Select NotaFiscal as NF, CodVarejo as Cod, QntProdutos as QNT, Conferir as Rest, Descricao as Descr From ConfereNFSaida where Varejista = '" + cboVarejista.Text + "' and Data = '" + Data + "' ";
            }

            cx.Conectar();
            SqlDataAdapter da = new SqlDataAdapter(sql, cx.c);
            DataSet ds = new DataSet();
            da.Fill(ds, "NotaFiscal");
            dgvConsulta.DataSource = ds.Tables["NotaFiscal"];
            cx.Desconectar();
            dgvConsulta.RowHeadersVisible = false;
            dgvConsulta.Columns["Descr"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvConsulta.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            dgvConsulta.AutoResizeColumns();
        }

        private void btnConcluir_Click(object sender, EventArgs e)
        {
            if(chbMultiplaQuantidade.Checked == false)
            {
                Concluir();
            }
            else
            {
                if (!int.TryParse(txtQtd.Text, out int qtdInformada))
                {
                    MessageBox.Show("Informe uma quantidade válida.");
                    txtQtd.Select();
                    txtQtd.SelectAll();
                    return;
                }

                if (!int.TryParse(txtSomaDasQuantidades.Text, out int somaQuantidades))
                {
                    MessageBox.Show("A soma das quantidades é inválida.");
                    return;
                }

                if (qtdInformada > somaQuantidades)
                {
                    MessageBox.Show("A quantidade informada é maior que a soma das quantidades disponíveis.");
                    return;
                }

                ConcluirEmMassa();
            }
            
        }

        public void Concluir()
        {
            // baixa a qnt da nota fiscal, se o flag estiver marcado
            consulta.comando = "";
            consulta.comando = "update ConfereNFSaida set Conferir = Conferir - 1 where idConfereNFSaida = (select top 1 idConfereNFSaida from ConfereNFSaida where CodVarejo = '" + txtCodigo.Text + "' and Conferir > 0 and Varejista = '" + cboVarejista.Text + "' and NotaFiscal in (" + notas + "))";
            consulta.Atualizar();
            consulta.PlayOK();
            AdcionarContador();
            lblUltimoColetado.Text = "ÚLTIMO: " + txtCodigo.Text + " - " + txtEanConsulta.Text + "\r\n" + txtDescricao.Text;
            btnLimpar.PerformClick();
            AtualizaContadores();
            ListarTudo(false);
            txtEAN.Select();
        }

        public void ConcluirEmMassa()
        {
            if (!int.TryParse(txtQtd.Text, out int qtdRestante) || qtdRestante <= 0)
            {
                consulta.PlayFail();
                MessageBox.Show("Informe uma quantidade válida para debitar.");
                txtQtd.Select();
                txtQtd.SelectAll();
                return;
            }

            string codigo = txtCodigo.Text;
            string varejista = cboVarejista.Text;

            SqlTransaction transacao = null;
            bool sucesso = false;            

            try
            {
                // 🔹 1️⃣ Abre conexão e inicia transação
                cx.Conectar();
                transacao = cx.c.BeginTransaction();

                SqlConnection conexao = cx.c;

                while (qtdRestante > 0)
                {
                    // 🔹 2️⃣ Busca próxima linha com saldo (usando sua função auxiliar)
                    if (!ConsultarProximaLinha(codigo, varejista, notas, conexao, transacao, out int id, out int conferirAtual))
                        break; // acabou o saldo

                    int debitar = Math.Min(qtdRestante, conferirAtual);

                    // 🔹 3️⃣ Executa o UPDATE dentro da transação
                    string sqlUpdate = @"
                UPDATE ConfereNFSaida 
                SET Conferir = Conferir - @Debitar 
                WHERE idConfereNFSaida = @Id";

                    using (SqlCommand cmdUpdate = new SqlCommand(sqlUpdate, conexao, transacao))
                    {
                        cmdUpdate.Parameters.AddWithValue("@Debitar", debitar);
                        cmdUpdate.Parameters.AddWithValue("@Id", id);
                        cmdUpdate.ExecuteNonQuery();
                    }

                    qtdRestante -= debitar;
                }

                // 🔹 4️⃣ Se tudo ocorreu bem, decide entre commit e aviso
                if (qtdRestante > 0)
                {
                    transacao.Rollback();
                    consulta.PlayFail();
                    MessageBox.Show($"Nem todas as unidades puderam ser debitadas. Faltaram {qtdRestante} para zerar o saldo.");
                }
                else
                {
                    transacao.Commit();                    
                    consulta.PlayOK();
                    MessageBox.Show("Débito realizado com sucesso!");
                    sucesso = true;
                }
                                
            }
            catch (Exception ex)
            {
                // ❌ Se algo der errado, desfaz tudo
                try { transacao?.Rollback(); } catch { }
                consulta.PlayFail();
                MessageBox.Show("Erro ao debitar quantidades:\n" + ex.Message);
            }
            finally
            {
                cx.Desconectar();
            }

            if (sucesso)
            {
                // 🔹 5️⃣ Atualiza a interface
                AdcionarContador();
                lblUltimoColetado.Text = "ÚLTIMO: " + txtCodigo.Text + " - " + txtEanConsulta.Text + "\r\n" + txtDescricao.Text;
                btnLimpar.PerformClick();
                AtualizaContadores();
                ListarTudo(false);
                txtEAN.Select();
            }
        }

        public bool ConsultarProximaLinha(string codigo, string varejista, string notas, SqlConnection conexao, SqlTransaction transacao, out int id, out int conferir)
        {
            id = 0;
            conferir = 0;

            try
            {
                string sql = $@"
            SELECT TOP 1 idConfereNFSaida, Conferir
            FROM ConfereNFSaida
            WHERE CodVarejo = '{codigo}'
              AND Varejista = '{varejista}'
              AND Conferir > 0
              AND NotaFiscal IN ({notas})
            ORDER BY idConfereNFSaida";

                using (SqlCommand cmd = new SqlCommand(sql, conexao, transacao))
                using (SqlDataReader dr = cmd.ExecuteReader())
                {
                    if (dr.Read())
                    {
                        id = Convert.ToInt32(dr["idConfereNFSaida"]);
                        conferir = Convert.ToInt32(dr["Conferir"]);
                        return true;
                    }
                }
            }
            catch (SqlException x)
            {
                MessageBox.Show("Falha em ConsultarProximaLinha:\n" + x.Message);
            }

            return false;
        }

        private void btbAdd_Click(object sender, EventArgs e)
        {

            if (lstColunas.Items.Contains(cboNotaFiscal.Text))
            {
                consulta.PlayFail();
                MessageBox.Show("ITEM JÁ CADASTRADO.");
            }
            else if (cboNotaFiscal.Text.Length == 0)
            {
                consulta.PlayFail();
                MessageBox.Show("SELECIONE O ITEM.");
            }
            else
            {
                lstColunas.Items.Add(cboNotaFiscal.Text);
                cboNotaFiscal.Text = null;
                cboNotaFiscal.Select();
                recebeColunas();
            }
        }

        public string notas = "";
        public void recebeColunas()
        {
            notas = "";
            int rows = lstColunas.Items.Count;
            foreach (string item in lstColunas.Items)
            {
                if (rows > 1)
                {
                    notas += "'" + item + "', ";
                }
                else
                {
                    notas += "'" + item + "' ";
                }
                rows--;
            }
            AtualizarListaNotas();
            ContarListados();
            limpar();
        }

        public void ContarListados()
        {
            int Listados = lstColunas.Items.Count;
            lblListados.Text = "NOTAS: " + Listados.ToString();
        }

        private void btnRem_Click(object sender, EventArgs e)
        {
            lstColunas.Items.Remove(lstColunas.SelectedItem);
            recebeColunas();
        }

        private void btnTodos_Click(object sender, EventArgs e)
        {
            lstColunas.Items.Clear();
            for (int i = 0; i < cboNotaFiscal.Items.Count; i++)
            {
                string value = cboNotaFiscal.GetItemText(cboNotaFiscal.Items[i]);
                lstColunas.Items.Add(value);
            }
            recebeColunas();
        }

        private void btnLimparLista_Click(object sender, EventArgs e)
        {
            lstColunas.Items.Clear();
            lblListados.Text = "NOTAS: 0";
        }

        public string Data = "";

        private void mtbDataNF_TextChanged(object sender, EventArgs e)
        {
            Data = "";
            if (mtbDataNF.Text.Length > 7)
            {
                Data = "";
                string check = "";
                string Vdata = "";
                ValidaData(mtbDataNF.Text, out check, out Vdata);
                if (check == "OK")
                {
                    Data = Vdata;

                    lstColunas.Items.Clear();
                    ContarVarejista();
                    PreencherCboNotaFiscal();
                    cboNotaFiscal.Select();
                }
                else
                {
                    mtbDataNF.SelectAll();
                }
            }
        }

        private void mtbDataNF_Leave(object sender, EventArgs e)
        {
            if (mtbDataNF.Text.Length == 0)
            {
                Data = "";
            }
            else if (mtbDataNF.Text.Length < 8)
            {
                consulta.PlayFail();
                MessageBox.Show("DATA INVÁLIDA");
                mtbDataNF.Focus();
            }
            else
            {
                string check = "";
                string Vdata = "";
                ValidaData(mtbDataNF.Text, out check, out Vdata);
                if (check == "OK")
                {
                    Data = Vdata;

                    lstColunas.Items.Clear();
                    ContarVarejista();
                    PreencherCboNotaFiscal();
                    cboNotaFiscal.Select();
                }
            }
        }

        public void ValidaData(string data, out string check, out string Vdata)
        {
            check = "";
            string Data = "";
            Vdata = data.Substring(0, 2) + "/" + data.Substring(2, 2) + "/" + data.Substring(4, 4);
            //MessageBox.Show(Vdata);
            try
            {
                DateTime teste = Convert.ToDateTime(Vdata);
                if (teste > DateTime.Now)
                {
                    consulta.PlayFail();
                    MessageBox.Show("DATA MAIOR QUE HOJE");
                    // mtbDataCompra.SelectAll();
                }
                else
                {
                    Data = Vdata;
                    // mtbDataCompra.Select();
                    check = "OK";
                }
            }
            catch (Exception)
            {
                consulta.PlayFail();
                MessageBox.Show("DIGITE UMA DATA VÁLIDA");
                // mtbDataCompra.SelectAll();
            }
        }

        private void btnListarNF_MouseEnter(object sender, EventArgs e)
        {
            //toolTip1.SetToolTip(btnListarNF, "CONSULTAR NFs DA LISTA.");
        }

        private void lnkListarPorData_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("DESEJA LISTAR AS NFs DO VAREJISTA NA DATA INFORMADA\r\n\r\nESSA CONSULTA RETORNA ATÉ OS ITENS JÁ CONFERIDOS? ", "Pergunta", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (cboVarejista.Text.Length > 0 || Data == "")
                {
                    ListarTudo(true);
                }
                else
                {
                    MessageBox.Show("INFORME O VAREJISTA E A DATA.");
                }
            }
        }

        public void AdcionarContador()
        {
            if (txtContagem.Lines.Length == 0)
            {
                txtContagem.Text = txtCodigo.Text;
            }
            else
            {
                txtContagem.Text = txtContagem.Text + "\r\n" + txtCodigo.Text;
            }
        }

        private void btnLimpaContador_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("DESEJA LIMPAR OS CÓDIGO JÁ COLETADOS? ", "Pergunta", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                txtContagem.Text = "";
            }
        }

        private void txtContagem_TextChanged(object sender, EventArgs e)
        {
            lblContagem.Text = txtContagem.Lines.Length.ToString();
        }

        private void chbMultiplaQuantidade_CheckedChanged(object sender, EventArgs e)
        {
            if (chbMultiplaQuantidade.Checked)
            {
                pnlMultiplasQuantidades.Visible = true;
            }
            else
            {
                pnlMultiplasQuantidades.Visible = false;
                txtSomaDasQuantidades.Text = "";
                txtQtd.Text = "";
            }
        }
    }
}
