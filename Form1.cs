using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Collections;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text.RegularExpressions;
using System.Net.Configuration;

namespace SolicitacaoCadastro
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

            textBoxUsuario.Text = "seuemail@seudominio.com.br";
            textBoxSenha.Text = "SenhaAqui.123";
            comboBoxNacional.Text = "SIM";
            comboBoxFornecedorExclusivo.Text = "NÃO";
        }

        private void buttonEnviar_Click(object sender, EventArgs e)
        {
            if (textBoxNome.Text == "" | textBoxEmail.Text == "" || textBoxDescricaoCompletaMaterial.Text == "" || textBoxDescricaoMaterialPortugues.Text == "" || textBoxDescricaoMaterialFabricante.Text == "" || comboBoxUtilizacao.Text == "")
            {
                MessageBox.Show("Existem dados obrigatórios não preenchido! Favor verificar");
            }
            else
            {
                try
                {
                    if(comboBoxFornecedorExclusivo.Text == "SIM" && textBoxJustificativaForExclusivo.Text == "")
                    {
                      
                            MessageBox.Show("Favor informar o motivo para fornecedor exclusivo!");
                            return;
                        
                    }
                    else
                    {
                        EnviarEmail(textBoxUsuario.Text, textBoxSenha.Text, textBoxNome.Text, textBoxEmail.Text, "cadastro@dominitoquerecebe.com.br");
                    }
                    DialogResult confirm = MessageBox.Show("Deseja limpar todos os campos?", "Salvarrr Arquivooo", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                    if (confirm.ToString().ToUpper() == "YES")
                    {
                        limparCampos();
                    }
                }
                catch
                {
                    MessageBox.Show("Ocorreu um erro interno. Favor procurar o TI.");
                }
                MessageBox.Show("Solicitação enviada com sucesso.");
            }
            
        }

        //ENVIAR MENSAGEM

        public void limparCampos() {

            textBoxDescricaoMaterialPortugues.Text = "";
            textBoxDescricaoMaterialFabricante.Text = "";
            textBoxCodigoFabricante.Text = "";
            textBoxCodigoWartsila.Text = "";
            textBoxNCM.Text = "";
            comboBoxNacional.Text = "SIM";
            comboBoxUtilizacao.Text = "";
            comboBoxFornecedorExclusivo.Text = "NÃO";
            textBoxJustificativaForExclusivo.Text = "";
            textBoxDescricaoCompletaMaterial.Text = "";
            textBoxAplicacaoProduto.Text = "";
        
        }

        public void EnviarEmail(string loginEmail, string senhaEmail, string nomeSolicitante, string emailSolicitante, string destinatario)
        {
            string corpoMensagem = "<p style='line-height: 100 %; font-size:20pt; font-family: verdana;'><b>Cadastro</b></p>" +
            "<p style = 'line-height: 100%; font-size:10pt; font-family: verdana;'><b> Solicitante:</b> " + textBoxNome.Text +" </br> " +
            "<b> E-mail:</b> "+ textBoxEmail.Text + "</br> </br> " +
            "<b> Descrição básica do material(PORTUGUÊS):</b > " + textBoxDescricaoMaterialPortugues.Text +" </br> " +
            "<b> Descrição Básica material(FABRICANTE):</b> "+ textBoxDescricaoMaterialFabricante.Text +" </br> " +
            "<b> Código fabricante:</b> " + textBoxCodigoFabricante.Text +" </br>" +
            "<b> Código Wartsila:</b> " + textBoxCodigoWartsila.Text +" </br>" +
            "<b> NCM:</b> " + textBoxNCM.Text +" </br>" +
            "<b> Nacional:</b> " + comboBoxNacional.Text +" </br>" +
            "<b> Utilização:</b> " + comboBoxUtilizacao.Text +" </br>" +
            "<b> Fornecedor exclusivo:</b> " + comboBoxFornecedorExclusivo.Text +" </br> </br>" +
            "<b> Justificativa do fornecedor exclusivo: </b> " + textBoxJustificativaForExclusivo.Text + " </br> " +
            "<b> Descrição completa do material:</b> " + textBoxDescricaoCompletaMaterial.Text +" </br> " +
            "<b> Aplicação do produto:</b> " + textBoxAplicacaoProduto.Text +" </br> </br>" +
            "<hr></br><img src = 'http://energeticasuape.com.br/cadastro/_imagens/logo-email.png'/></p> ";

            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            client.Host = "outlook.office365.com";
            //SmtpClient client = new SmtpClient("email-ssl.com.br", 587);
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(loginEmail, senhaEmail);
            MailMessage mail = new MailMessage();
            mail.Sender = new System.Net.Mail.MailAddress(loginEmail, "Work Flow");
            mail.From = new MailAddress(loginEmail, "Work Flow");
            mail.To.Add(new MailAddress(destinatario, "Cadastro | Suape Energia"));
            mail.CC.Add(new MailAddress(emailSolicitante, nomeSolicitante));
            mail.Subject = "Solicitação de Cadastro feita por " + nomeSolicitante;
            //mail.Body = " Nova solicitação de cadastro:<br/> Nome:  " + nomeRemetente + "<br/> Email Solicitante : " + remetente + " <br/> Mensagem : " + mensagem;
            mail.Body = corpoMensagem;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;
            try
            {
                client.Send(mail);
            }
            catch (System.Exception erro)
            {
                //MessageBox.Show(erro);
            }
            finally
            {
                mail = null;
            }
        }
        

        /// <summary>
        /// Confirma a validade de um email
        /// </summary>
        /// <param name="enderecoEmail">Email a ser validado</param>
        /// <returns>Retorna True se o email for valido</returns>
        public static bool ValidaEnderecoEmail(string enderecoEmail)
        {
            try
            {
                //define a expressão regulara para validar o email
                string texto_Validar = enderecoEmail;
                Regex expressaoRegex = new Regex(@"\w+@[a-zA-Z_]+?\.[a-zA-Z]{2,3}");

                // testa o email com a expressão
                if (expressaoRegex.IsMatch(texto_Validar))
                {
                    // o email é valido
                    return true;
                }
                else
                {
                    // o email é inválido
                    return false;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void comboBoxNacional_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            this.Close();
        }
    }
}
