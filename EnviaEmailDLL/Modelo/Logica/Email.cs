using EnviaEmailDLL.Modelo.Dominio;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;

namespace EnviaEmailDLL.Modelo.Logica
{
    /// <summary>
    /// Classe para tratar o envio de email.
    /// </summary>
    [ComVisible(true)]
    [ProgId("EnviaEmailDLL.Email")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Email
    {
        private const bool EmailNaoEnviado = false;
        private List<string> _destinatarios = new List<string>();
        private List<string> _emailsBcc = new List<string>();
        private List<string> _emailsCc = new List<string>();
        private List<string> _anexos = new List<string>();

        /// <summary>
        /// Propriedade criado para utilização dentro do VFP.
        /// * Recebe a instancia de EmailConfiguracao.
        ///     - OBRIGATÓRIO SETAR ESTÁ PROPRIEDADE DENTRO DO VFP ANTES DO METODO "Enviar".
        /// </summary>
        public EmailConfiguracao EmailConfiguracao { get; set; }

        public Email() { }

        /// <summary>
        /// Realiza o envio de email.
        /// </summary>
        /// <param name="assunto">Assunto do email.</param>
        /// <param name="emailDestinatarios">Array de string com os emails de destino.</param>
        /// <param name="mensagem">Mensagem do corpo do email.</param>
        /// <param name="anexos">Array de string com o caminho completo dos arquivos para anexar no email.</param>
        /// <param name="emailsBcc">Array de string com os emails para copia oculta.</param>
        /// <param name="emailsCC">Array de string com os emails para copia.</param>
        /// <param name="isBodyHtml">Booleano para definir se a mensagem será tratada como HTML.</param>
        /// <param name="prioridade">Integer para definir a prioridade do email.<para>Normal = 0, Low = 1, High = 2</para></param>
        /// <param name="deliveryNotification">Integer para definir o DeliveryNotification.<para>None = 0, OnSuccess = 1, OnFailure = 2, Delay = 4, Never = 134217728</para></param>
        /// <returns>Booleano</returns>
        public bool Enviar(string assunto, string[] emailDestinatarios, string mensagem, string[] anexos,
            string[] emailsBcc, string[] emailsCC, bool isBodyHtml = false, int prioridade = 0, int deliveryNotification = 0)
        {
            try
            {
                using (var mailMessage = new MailMessage())
                {
                    var mailAddress = new MailAddress(
                        EmailConfiguracao.EmailRemetente.Trim(),
                        EmailConfiguracao.NomeRemetente.Trim());

                    mailMessage.Priority = (MailPriority)prioridade;
                    mailMessage.DeliveryNotificationOptions = (DeliveryNotificationOptions)deliveryNotification;
                    mailMessage.IsBodyHtml = isBodyHtml;
                    mailMessage.From = mailAddress;
                    mailMessage.Subject = assunto.Trim();

                    // Popula list de emails destinatario
                    PopulaMailAddressCollection(mailMessage.To, emailDestinatarios);

                    // Popula list de emails BCC
                    PopulaMailAddressCollection(mailMessage.Bcc, emailsBcc);

                    // Popula list de emails CC
                    PopulaMailAddressCollection(mailMessage.CC, emailsCC);

                    mailMessage.Body = mensagem.Trim();

                    if (anexos?.Length > 0)
                    {
                        foreach (var anexo in anexos)
                            mailMessage.Attachments.Add(new Attachment(anexo.Trim()));
                    }

                    ConfiguraSmtpEnviaEmail(mailMessage);

                    return !EmailNaoEnviado;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        //        
        /// <summary>
        /// Método criado para utilização dentro do VFP.
        /// * OBS.:
        ///    - Obrigatório chamar o método "InstanciaEmailConfiguracao" e "AdicionaEmailDestinatarios" ANTES de usar este método.
        /// </summary>
        /// <param name="assunto">Assunto do email.</param>
        /// <param name="mensagem">Mensagem do corpo do email.</param>
        /// <param name="isBodyHtml">Booleano para definir se a mensagem será tratada como HTML.</param>
        /// <param name="prioridade">Integer para definir a prioridade do email.<para>Normal = 0, Low = 1, High = 2</para></param>
        /// <param name="deliveryNotification">Integer para definir o DeliveryNotification.<para>None = 0, OnSuccess = 1, OnFailure = 2, Delay = 4, Never = 134217728</para></param>
        /// <returns>Booleano</returns>
        public bool Enviar(string assunto, string mensagem, bool isBodyHtml = false,
            int prioridade = 0, int deliveryNotification = 0)
        {
            try
            {
                using (var mailMessage = new MailMessage())
                {
                    var mailAddress = new MailAddress(
                        EmailConfiguracao.EmailRemetente.Trim(),
                        EmailConfiguracao.NomeRemetente.Trim());

                    mailMessage.Priority = (MailPriority)prioridade;
                    mailMessage.DeliveryNotificationOptions = (DeliveryNotificationOptions)deliveryNotification;
                    mailMessage.IsBodyHtml = isBodyHtml;
                    mailMessage.From = mailAddress;
                    mailMessage.Subject = assunto.Trim();

                    // Popula list de emails destinatario
                    PopulaMailAddressCollection(mailMessage.To, _destinatarios);

                    // Popula list de emails BCC
                    PopulaMailAddressCollection(mailMessage.Bcc, _emailsBcc);

                    // Popula list de emails CC
                    PopulaMailAddressCollection(mailMessage.CC, _emailsCc);

                    mailMessage.Body = mensagem.Trim();

                    if (_anexos?.Count > 0)
                    {
                        foreach (var anexo in _anexos)
                            mailMessage.Attachments.Add(new Attachment(anexo.Trim()));
                    }

                    ConfiguraSmtpEnviaEmail(mailMessage);

                    return !EmailNaoEnviado;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Método criado para utilização dentro do VFP.
        /// * Popula o List _destinatarios para ser utilizado pelo método "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        ///     - OBRIGATÓRIO CHAMAR ESTE MÉTODO DENTRO DO VFP ANTES DO "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        /// </summary>
        /// <param name="emailDestinatario"></param>
        public void AdicionaEmailDestinatarios(string emailDestinatario) => _destinatarios.Add(emailDestinatario);

        /// <summary>
        /// Método criado para utilização dentro do VFP.
        /// * Popula o List _emailsBcc para ser utilizado pelo método "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        ///     - OPCIONAL CHAMAR ESTE MÉTODO DENTRO DO VFP ANTES DO "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        /// </summary>
        /// <param name="email"></param>
        public void AdicionaEmailBcc(string email) => _emailsBcc.Add(email);

        /// <summary>
        /// Método criado para utilização dentro do VFP.
        /// * Popula o List _emailsCc para ser utilizado pelo método "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        ///     - OPCIONAL CHAMAR ESTE MÉTODO DENTRO DO VFP ANTES DO "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        /// </summary>
        /// <param name="email"></param>
        public void AdicionaEmailCc(string email) => _emailsCc.Add(email);

        /// <summary>
        /// Método criado para utilização dentro do VFP.
        /// * Popula o List _anexos para ser utilizado pelo método "Enviar(string pAssunto, string pMensagem, bool pIsBodyHtml = false, int pPrioridade = 0)".
        ///     - CHAMADA AO MÉTODO DENTRO DO VFP É OPCIONAL.
        /// </summary>
        /// <param name="caminhoAnexo"></param>
        public void AdicionaAnexos(string caminhoAnexo) => _anexos.Add(caminhoAnexo);
            
        private void ConfiguraSmtpEnviaEmail(MailMessage mailMsg)
        {
            try
            {
                using (var smptClient = new SmtpClient(EmailConfiguracao.ServidorSmtp, EmailConfiguracao.PortaSmtp))
                {
                    smptClient.UseDefaultCredentials = false;
                    smptClient.EnableSsl = EmailConfiguracao.AutenticaSSL;

                    if (EmailConfiguracao.Autentica)
                        smptClient.Credentials = new NetworkCredential(EmailConfiguracao.LoginMail, EmailConfiguracao.SenhaMail);

                    smptClient.Send(mailMsg);
                }
            }
            catch (ArgumentException ex)
            {
                throw new Exception(ex.Message);
            }
            catch (SmtpException ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void PopulaMailAddressCollection(MailAddressCollection mailAddressCollection, IEnumerable<string> emails)
        {
            try
            {
                foreach (var email in emails)
                    mailAddressCollection.Add(email.Trim());
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}