using EnviaEmailDLL.Modelo.Dominio;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EnviaEmailDLL.Modelo.Logica.Tests
{
    [TestClass]
    public class EmailTeste
    {
        private Email _email;
        private Configuracao _config;
        private string[] _anexos;
        private string[] _destinatarios;
        private string[] _emailsBcc;
        private string[] _emailsCc;

        [TestInitialize]
        public void SetUp()
        {
            _email = new Email();
            _config = new Configuracao();

            _anexos = new string[]
            {
                @"anexo1.pdf",
                @"anexo2.pdf"
            };

            _destinatarios = new string[]
            {
                @"seuemail@spaceinformatica.com.br",
            };

            _emailsBcc = new string[]
            {
                @"seuemail@hotmail.com"
            };

            _emailsCc = new string[]
            {
                @"seuemail@gmail.com"
            };
        }

        [TestMethod]
        public void EnviaEmailSmtpUOL()
        {
            var email = @"seuemail@spaceinformatica.com.br";
            var emailConfig = new EmailConfiguracao()
            {
                NomeRemetente = "REMETENTE UOL",
                EmailRemetente = email,
                ServidorSmtp = @"smtp.spaceinformatica.com.br",
                PortaSmtp = 587,
                LoginMail = email,
                SenhaMail = @"senha",
                Autentica = true,
                AutenticaSSL = false
            };

            //_config.SalvaConfiguracoesEmail(emailConfig);

            _email.EmailConfiguracao = emailConfig;

            var result = _email.Enviar(
                "ASSUNTO TESTE UOL",
                _destinatarios,
                "<h1>MENSAGEM TESTE UOL</h1>",
                _anexos,
                _emailsBcc,
                _emailsCc,
                true);

            Assert.IsTrue(result);
        }

        [TestMethod]
        [Description("Necessário Permitir Apps menos seguros nas configurações do email.")]
        public void EnviaEmailSmtpGmail()
        {
            var email = @"seuemail@gmail.com";
            var emailConfig = new EmailConfiguracao()
            {
                NomeRemetente = "REMETENTE GMAIL",
                EmailRemetente = email,
                ServidorSmtp = @"smtp.gmail.com",
                PortaSmtp = 587,
                LoginMail = email,
                SenhaMail = @"senha",
                Autentica = true,
                AutenticaSSL = true
            };

            //_config.SalvaConfiguracoesEmail(emailConfig);

            _email.EmailConfiguracao = emailConfig;

            var result = _email.Enviar(
                "ASSUNTO TESTE GMAIL",
                _destinatarios,
                "MENSAGEM TESTE GMAIL",
                _anexos,
                _emailsBcc,
                _emailsCc);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EnviaEmailSmtpHomail()
        {
            var email = @"seuemail@hotmail.com";
            var emailConfig = new EmailConfiguracao()
            {
                NomeRemetente = "REMETENTE HOTMAIL",
                EmailRemetente = email,
                ServidorSmtp = @"smtp.live.com",
                PortaSmtp = 587,
                LoginMail = email,
                SenhaMail = @"senha",
                Autentica = true,
                AutenticaSSL = true
            };

            //_config.SalvaConfiguracoesEmail(emailConfig);

            _email.EmailConfiguracao = emailConfig;

            var result = _email.Enviar(
                "ASSUNTO TESTE HOMAIL",
                _destinatarios,
                "MENSAGEM TESTE HOMAIL",
                _anexos,
                _emailsBcc,
                _emailsCc);

            Assert.IsTrue(result);
        }

        [TestMethod]
        [Description("Necessário Permitir Apps menos seguros nas configurações do email.")]
        public void EnviaEmailSmtpYahoo()
        {
            var email = @"seuemail@yahoo.com.br";
            var emailConfig = new EmailConfiguracao()
            {
                NomeRemetente = "REMETENTE YAHOO",
                EmailRemetente = email,
                ServidorSmtp = @"smtp.mail.yahoo.com",
                PortaSmtp = 587,
                LoginMail = email,
                SenhaMail = @"senha",
                Autentica = true,
                AutenticaSSL = true
            };

            //_config.SalvaConfiguracoesEmail(emailConfig);

            _email.EmailConfiguracao = emailConfig;

            var result = _email.Enviar(
                "ASSUNTO TESTE YAHOO",
                _destinatarios,
                "MENSAGEM TESTE YAHOO",
                _anexos,
                _emailsBcc,
                _emailsCc);

            Assert.IsTrue(result);
        }

        [TestMethod]
        [Description("Usa a sobrecarga do método Enviar")]
        public void EnviarEmailTeste()
        {
            var email = @"seuemail@spaceinformatica.com.br";
            var emailConfig = new EmailConfiguracao()
            {
                NomeRemetente = "REMETENTE UOL",
                EmailRemetente = email,
                ServidorSmtp = @"smtp.spaceinformatica.com.br",
                PortaSmtp = 587,
                LoginMail = email,
                SenhaMail = @"senha",
                Autentica = true,
                AutenticaSSL = false
            };

            //_config.SalvaConfiguracoesEmail(emailConfig);

            _email.EmailConfiguracao = emailConfig;

            foreach (var dest in _destinatarios)
                _email.AdicionaEmailDestinatarios(dest);

            foreach (var bcc in _emailsBcc)
                _email.AdicionaEmailBcc(bcc);

            foreach (var cc in _emailsCc)
                _email.AdicionaEmailCc(cc);

            foreach (var anx in _anexos)
                _email.AdicionaAnexos(anx);

            var result = _email.Enviar("ASSUNTO TESTE UOL", "MENSAGEM TESTE UOL", false, 2);

            Assert.IsTrue(result);
        }
    }
}
