using Entity;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Documento.Model;

namespace Contrato_Basic
{
    public partial class formContrato : Form
    {
        public formContrato()
        {
            InitializeComponent();
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void bt_Salvar_Click(object sender, EventArgs e)
        {
            Pessoa locador = new Pessoa()
            {
                Nome = txtNomeLocador.Text,
                RG = mskRGlocador.Text,
                CPF = mskCPFLocador.Text,
                Profissao = txtProfLocador.Text,
                IsMasculino = chkMascLocador.Checked,
                Endereco = new Endereco()
                {
                    Rua = txtRuaLocador.Text,
                    Bairro = txtBairroLocador.Text,
                    Cidade = txtCidadeLocador.Text,
                    Estado = txtEstadoLocador.Text,
                },
                Pagamento = new Pagamento()
                {
                    Banco = txtBancoLocador.Text,
                    Agencia = txtAgenciaLocador.Text,
                    Conta = txtContaLocador.Text,
                    PIX = txtPIXLocador.Text,
                }
            };

            Pessoa locatario = new Pessoa()
            {
                Nome = txtNomeLocatario.Text,
                RG = mskRGLocatario.Text,
                CPF = mskCPFLocatario.Text,
                Profissao = txtProfLocatario.Text,
                IsMasculino = chkMascLocatario.Checked,
                Endereco = new Endereco()
                {
                    Rua = txtRuaLocador.Text,
                    Bairro = txtBairroLocatario.Text,
                    Cidade = txtCidadeLocatario.Text,
                    Estado = txtEstadoLocatario.Text,
                    Periodo = Convert.ToInt32(txtPeriodoContrato.Text),
                    InicioLocacao = dtpInicio.Value,
                    TerminoLocacao = dtpInicio.Value.AddMonths(Convert.ToInt32(txtPeriodoContrato.Text)),
                    Vencimento = Convert.ToInt32(numericUpDownVencimento.Value),
                    Valor = txtValorContrato.Text,
                },
                Pagamento = new Pagamento(),
            };

            Documento.Model.Documento documento = new Documento.Model.Documento(locador, locatario);
        }
    }
}
