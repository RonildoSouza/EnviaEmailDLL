using EnviaEmailDLL.Modelo.Dominio;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EnviaEmailDLL.Modelo.Logica.Tests
{
    [TestClass]
    public class ConfiguracaoTeste
    {
        private Configuracao _config;

        [TestInitialize]
        public void SetUp()
        {
            _config = new Configuracao();
        }

        [TestMethod]
        public void _1_SalvaConfiguracoesEmailTeste()
        {
            var emailConfig = new EmailConfiguracao()
            {
                NomeRemetente = "REMETENTE ALTERADO NO TESTE UNITÁRIO",
                EmailRemetente = "EmailTesteUnitario@email.com",
                ServidorSmtp = @"smtp.TesteUnitario.com.br",
                PortaSmtp = 587,
                LoginMail = "EmailLoginTesteUnitario@email.com",
                SenhaMail = @"SenhaTesteUnitario",
                Autentica = true,
                AutenticaSSL = false
            };

            var result = _config.SalvaConfiguracoesEmail(emailConfig);

            Assert.IsTrue(result);
        }

        [TestMethod]
        public void _2_RetornaConfiguracoesEmailXMLTeste()
        {
            var emailConfig = _config.RetornaConfiguracoesEmail();

            Assert.IsTrue(emailConfig != null);
        }
    }
}