using System.Runtime.InteropServices;
using System.Xml.Serialization;

namespace EnviaEmailDLL.Modelo.Dominio
{
    /// <summary>
    /// Classe modelo para facilitar a manipulação das configurações de email.
    /// </summary>
    [ComVisible(true)]
    [ProgId("EnviaEmailDLL.EmailConfiguracao")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [XmlRoot(nameof(EmailConfiguracao))]
    public class EmailConfiguracao
    {
        [XmlElement]
        public string NomeRemetente { get; set; }
        [XmlElement]
        public string EmailRemetente { get; set; }
        [XmlElement]
        public string ServidorSmtp { get; set; }
        [XmlElement]
        public int PortaSmtp { get; set; }
        [XmlElement]
        public string LoginMail { get; set; }
        [XmlElement]
        public string SenhaMail { get; set; }
        [XmlElement]
        public bool Autentica { get; set; }
        [XmlElement]
        public bool AutenticaSSL { get; set; }
    }
}
