using EnviaEmailDLL.Helper;
using EnviaEmailDLL.Modelo.Dominio;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
//using Newtonsoft.Json;

namespace EnviaEmailDLL.Modelo.Logica
{
    /// <summary>
    /// Classe responsável por manipular o arquivo de configurações para envio de email.
    /// </summary>
    [ComVisible(true)]
    [ProgId("EnviaEmailDLL.Configuracao")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Configuracao
    {
        private const string ArquivoJson = @"EmailConfiguracao.json";
        private const bool ArquivoNaoSalvo = false;
        private const string DeserializeRootElementName = nameof(EmailConfiguracao);

        /// <summary>
        /// Retorna as configurações definidas no arquivo JSON.
        /// </summary>
        /// <returns>Objeto do tipo EmailConfiguracao</returns>
        public EmailConfiguracao RetornaConfiguracoesEmail()
        {
            try
            {
                var emailConfig = DeserializaEmailConfiguracaoJson();
                return emailConfig;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Salva as configurações de email no arquivo JSON.
        /// </summary>        
        /// <param name="emailConfig">Objeto do tipo EmailConfiguracao</param>
        /// <returns>Booleano</returns>
        /// ## Estrutura do arquivo JSON.
        ///    > {
        ///    >    "NomeRemetente": "REMETENTE",
        ///    >    "EmailRemetente": "seuemail@email.com",
        ///    >    "ServidorSmtp": "smtp.seuprovedor.com",
        ///    >    "PortaSmtp": 587,
        ///    >    "LoginMail": "seuemail@email.com",
        ///    >    "SenhaMail": "pexfgEuZuLDVvXWsHl5gww==",
        ///    >    "Autentica": true,
        ///    >    "AutenticaSSL": true
        ///    > }
        /// 
        /// ### OBS.: 
        ///    * Se o arquivo JSON não existir o mesmo será criado.
        public bool SalvaConfiguracoesEmail(EmailConfiguracao emailConfig)
        {
            try
            {
                SerializaEmailConfiguracaoJson(emailConfig);
                return !ArquivoNaoSalvo;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void SerializaEmailConfiguracaoJson(EmailConfiguracao emailConfig)
        {
            emailConfig.SenhaMail = Encryption64.Encrypt(emailConfig.SenhaMail);
            var emailCfg = InvokeSerializeObject(emailConfig);
            #region Implementação para utilização da DLL Newtonsoft.Json.dll como referência
            //var emailConfig = JsonConvert.SerializeObject(pEmailConfig, Newtonsoft.Json.Formatting.Indented);
            #endregion

            if (!File.Exists(ArquivoJson))
                File.Create(ArquivoJson).Close();

            File.WriteAllText(ArquivoJson, emailCfg);
        }

        private EmailConfiguracao DeserializaEmailConfiguracaoJson()
        {
            EmailConfiguracao emailConfig = null;

            if (File.Exists(ArquivoJson))
            {
                #region Implementação para utilização da DLL Newtonsoft.Json.dll como referência
                //emailConfig = JsonConvert
                //    .DeserializeObject<EmailConfiguracao>(File.ReadAllText(ArquivoJson));
                #endregion
                emailConfig = InvokeDeserializeObject(File.ReadAllText(ArquivoJson));
                emailConfig.SenhaMail = Encryption64.Decrypt(emailConfig.SenhaMail);
            }

            return emailConfig;
        }

        #region Metodos criados para não precisar referenciar a DLL Newtonsoft.Json.dll.
        private Type GetTypeJsonConvert()
        {
            Type type = null;
            const string NewtonsoftJsonDll = "Newtonsoft.Json.dll";

            try
            {
                Assembly assembly = Assembly.LoadFrom(NewtonsoftJsonDll);
                type = assembly.GetType("Newtonsoft.Json.JsonConvert");

                return type;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private EmailConfiguracao InvokeDeserializeObject(string json)
        {
            Type type = GetTypeJsonConvert();

            var metodo = type.GetMethods()
                .Where(m => m.Name == "DeserializeObject" && m.IsGenericMethod == true)
                .FirstOrDefault();

            var metodosParam = new object[] { json };

            var result = metodo
                .MakeGenericMethod(new Type[] { new EmailConfiguracao().GetType() })
                .Invoke(type, metodosParam);

            return result as EmailConfiguracao;
        }

        private string InvokeSerializeObject(EmailConfiguracao emailConfig)
        {
            int indentar = 1; // Valores do Enum Newtonsoft.Json.JsonConvert.Formatting são 1 e 0.
            Type type = GetTypeJsonConvert();

            var metodo = type.GetMethods()
                .Where(m => m.Name == "SerializeObject" && m.ReturnType == typeof(string))
                .ToArray()[1];

            var metodosParam = new object[] { emailConfig, indentar };

            var result = metodo.Invoke(type, metodosParam);

            return result as string;
        }
        #endregion
    }
}
