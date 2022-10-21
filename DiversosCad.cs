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
    public partial class DiversosCad : Form
    {
        public DiversosCad()
        {
            InitializeComponent();
        }
        string corpoMensagem;
        string mensagemCompleta;
        private void buttonEnviar_Click(object sender, EventArgs e)
        {
            if(textBoxNome.Text == "" | textBoxEmail.Text == "" | textBoxAplicacaoProduto.Text == "")
            {
                MessageBox.Show("Favor preencher os campos obrigatórios");
                return;
            }
            else
            {              
                try
                {
                    MessageBox.Show("Favor clicar em OK e aguardar a próxima mensagem!");

                    EnviarEmail("seuemail@seudominio.com.br", "SenhaAqui.123", textBoxNome.Text, textBoxEmail.Text, "cadastro@dominiododestinatario.com.br");          
                    
                    MessageBox.Show("E-mail enviado com sucesso!!");

                    dataGridView1.Rows.Clear();
                    textBox1.Text = "";
                    textBoxAplicacaoProduto.Text = "";

                }
                catch(Exception ex)
                {
                    MessageBox.Show("Erro no envio do e-mail! Erro: " + ex);
                }
            }    
        }

        public void EnviarEmail(string loginEmail, string senhaEmail, string nomeSolicitante, string emailSolicitante, string destinatario)
        {

            string cabecalhoMensagem = "<p style='line-height: 100 %; font-size:20pt; font-family: verdana;'><b>Cadastro</b></p>" +
           "<p style = 'line-height: 100%; font-size:10pt; font-family: verdana;'><b> Solicitante:</b> " + textBoxNome.Text + " </br> " +
           "<b> E-mail:</b> " + textBoxEmail.Text + "</br> </br> " +
           "<b> Códigos e produtos para cadastrar: </b > </br></br>" +
           "<table style='width:75%; border:1px solid gray;'>" +
           "<tr style='border:1px solid gray;'>" +
           "<th style='width:50px; border:1px solid gray;'>COD PRODUTO</th>" +
           "<th style='width:150px; border:1px solid gray;'>DESCRIÇÃO</th>" +
           "</tr>";

            string finalMensagem = "</table>" +
                "</br> <b>Aplicação:</b > " + textBoxAplicacaoProduto.Text + " </br></br> " +
           "<b> Observações:</b> " + textBox1.Text + " </br> " +
           "<hr></br><img src = 'http://seudominio.com.br/cadastro/_imagens/logo-email.png'/></p> ";

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //if(Convert.ToString(dataGridView1.Rows[i].Cells[0].Value) == "")
                //{
                    //break;
                //}
                //else {
                    corpoMensagem = corpoMensagem + "<tr style='border:1px solid gray;'>" +
                 "<td style='border:1px solid gray;'>" + Convert.ToString(dataGridView1.Rows[i].Cells[0].Value) + "</td>" +
                 "<td style='border:1px solid gray;'>" + Convert.ToString(dataGridView1.Rows[i].Cells[1].Value) + "</td >" +
                  "</tr>";
                //}
                
            }

            mensagemCompleta = cabecalhoMensagem + corpoMensagem + finalMensagem;

            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            client.Host = "outlook.office365.com";
            //SmtpClient client = new SmtpClient("email-ssl.com.br", 587);
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential(loginEmail, senhaEmail);
            MailMessage mail = new MailMessage();
            mail.Sender = new System.Net.Mail.MailAddress(loginEmail, "Work Flow");
            mail.From = new MailAddress(loginEmail, "Work Flow");
            mail.To.Add(new MailAddress(destinatario, "Cadastro | Sua Empresa"));
            mail.CC.Add(new MailAddress(emailSolicitante, nomeSolicitante));
            mail.Subject = "Solicitação de Cadastro feita por " + nomeSolicitante;
            //mail.Body = " Nova solicitação de cadastro:<br/> Nome:  " + nomeRemetente + "<br/> Email Solicitante : " + remetente + " <br/> Mensagem : " + mensagem;
            mail.Body = mensagemCompleta;
            mail.IsBodyHtml = true;
            mail.Priority = MailPriority.High;

            corpoMensagem = "";

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
    }
}
