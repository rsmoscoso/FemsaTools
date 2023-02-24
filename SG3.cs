using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;

namespace SG3ServiceReference
{
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.3")]
    [System.ServiceModel.ServiceContractAttribute(Namespace = "https://sg3.executiva.adm.br/femsa/?r=webservice/soap/index", ConfigurationName = "ServiceReference.ExecutivaPort")]
    public interface ExecutivaPort
    {

        [System.ServiceModel.OperationContractAttribute(Action = "https://sg3.executiva.adm.br/femsa/?r=webservice/soap/index#getLiberacao", ReplyAction = "*")]
        [System.ServiceModel.XmlSerializerFormatAttribute(SupportFaults = true)]
        System.Threading.Tasks.Task<SG3ServiceReference.getLiberacaoResponse> getLiberacaoAsync(SG3ServiceReference.getLiberacaoRequest request);
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.3")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace = "https://sg3.executiva.adm.br/femsa/?r=webservice/soap/index")]
    public partial class Liberacao
    {

        private string id_alocacaoField;

        private string id_estabelecimentoField;

        private string id_empresa_terceiraField;

        private string id_colaboradorField;

        private string estabelecimentoField; 

        private string empresa_terceiraField;

        private string colaboradorField;

        private string rgField;

        private string cpfField;

        private string pisField;

        private string tipo_terceiroField;

        private string funcaoField;

        private string data_inicioField;

        private string data_fimField;

        private string data_integracaoField;

        private string data_vencimento_integracaoField;

        private string data_demissaoField;

        private string liberadoField;
        private string crachafield;
        private string status_alocacaofield;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string id_alocacao
        {
            get
            {
                return this.id_alocacaoField;
            }
            set
            {
                this.id_alocacaoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string id_estabelecimento
        {
            get
            {
                return this.id_estabelecimentoField;
            }
            set
            {
                this.id_estabelecimentoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string id_empresa_terceira
        {
            get
            {
                return this.id_empresa_terceiraField;
            }
            set
            {
                this.id_empresa_terceiraField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string id_colaborador
        {
            get
            {
                return this.id_colaboradorField;
            }
            set
            {
                this.id_colaboradorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string estabelecimento
        {
            get
            {
                return this.estabelecimentoField;
            }
            set
            {
                this.estabelecimentoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string empresa_terceira
        {
            get
            {
                return this.empresa_terceiraField;
            }
            set
            {
                this.empresa_terceiraField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string colaborador
        {
            get
            {
                return this.colaboradorField;
            }
            set
            {
                this.colaboradorField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string rg
        {
            get
            {
                return this.rgField;
            }
            set
            {
                this.rgField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string cpf
        {
            get
            {
                return this.cpfField;
            }
            set
            {
                this.cpfField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string pis
        {
            get
            {
                return this.pisField;
            }
            set
            {
                this.pisField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string tipo_terceiro
        {
            get
            {
                return this.tipo_terceiroField;
            }
            set
            {
                this.tipo_terceiroField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string funcao
        {
            get
            {
                return this.funcaoField;
            }
            set
            {
                this.funcaoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string data_inicio
        {
            get
            {
                return this.data_inicioField;
            }
            set
            {
                this.data_inicioField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string data_fim
        {
            get
            {
                return this.data_fimField;
            }
            set
            {
                this.data_fimField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string data_integracao
        {
            get
            {
                return this.data_integracaoField;
            }
            set
            {
                this.data_integracaoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string data_vencimento_integracao
        {
            get
            {
                return this.data_vencimento_integracaoField;
            }
            set
            {
                this.data_vencimento_integracaoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string data_demissao
        {
            get
            {
                return this.data_demissaoField;
            }
            set
            {
                this.data_demissaoField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string liberado
        {
            get
            {
                return this.liberadoField;
            }
            set
            {
                this.liberadoField = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string cracha
        {
            get
            {
                return this.crachafield;
            }
            set
            {
                this.crachafield = value;
            }
        }

        [System.Xml.Serialization.XmlElementAttribute(Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string status_alocacao
        {
            get
            {
                return this.status_alocacaofield;
            }
            set
            {
                this.status_alocacaofield = value;
            }
        }
    }

    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.3")]
    [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
    [System.ServiceModel.MessageContractAttribute(WrapperName = "getLiberacao", WrapperNamespace = "https://sg3.executiva.adm.br/femsa/?r=webservice/soap/index", IsWrapped = true)]
    public partial class getLiberacaoRequest
    {

        [System.ServiceModel.MessageBodyMemberAttribute(Namespace = "", Order = 0)]
        public SG3ServiceReference.Liberacao liberacao;

        [System.ServiceModel.MessageBodyMemberAttribute(Namespace = "", Order = 1)]
        public int tipo_registro;

        public getLiberacaoRequest()
        {
        }

        public getLiberacaoRequest(SG3ServiceReference.Liberacao liberacao, int tipo_registro)
        {
            this.liberacao = liberacao;
            this.tipo_registro = tipo_registro;
        }
    }

    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.3")]
    [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
    [System.ServiceModel.MessageContractAttribute(WrapperName = "getLiberacaoResponse", WrapperNamespace = "https://sg3.executiva.adm.br/femsa/?r=webservice/soap/index", IsWrapped = true)]
    public partial class getLiberacaoResponse
    {

        [System.ServiceModel.MessageBodyMemberAttribute(Namespace = "", Order = 0)]
        [System.Xml.Serialization.XmlArrayAttribute()]
        [System.Xml.Serialization.XmlArrayItemAttribute("item", Form = System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable = false)]
        public SG3ServiceReference.Liberacao[] @return;

        public getLiberacaoResponse()
        {
        }

        public getLiberacaoResponse(SG3ServiceReference.Liberacao[] @return)
        {
            this.@return = @return;
        }
    }

    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.3")]
    public interface ExecutivaPortChannel : SG3ServiceReference.ExecutivaPort, System.ServiceModel.IClientChannel
    {
    }

    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.Tools.ServiceModel.Svcutil", "2.0.3")]
    public partial class ExecutivaPortClient : System.ServiceModel.ClientBase<SG3ServiceReference.ExecutivaPort>, SG3ServiceReference.ExecutivaPort
    {

        /// <summary>
        /// Implemente este método parcial para configurar o ponto de extremidade de serviço.
        /// </summary>
        /// <param name="serviceEndpoint">O ponto de extremidade a ser configurado</param>
        /// <param name="clientCredentials">As credenciais do cliente</param>
        static partial void ConfigureEndpoint(System.ServiceModel.Description.ServiceEndpoint serviceEndpoint, System.ServiceModel.Description.ClientCredentials clientCredentials);

        public ExecutivaPortClient() :
                base(ExecutivaPortClient.GetDefaultBinding(), ExecutivaPortClient.GetDefaultEndpointAddress())
        {

            this.Endpoint.Name = EndpointConfiguration.ExecutivaPort.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }

        public ExecutivaPortClient(System.ServiceModel.Description.ClientCredentials clientCredentials) :
        base(ExecutivaPortClient.GetDefaultBinding(), ExecutivaPortClient.GetDefaultEndpointAddress())
        {
            this.Endpoint.Name = EndpointConfiguration.ExecutivaPort.ToString();
            ConfigureEndpoint(this.Endpoint, clientCredentials);
        }

        public ExecutivaPortClient(EndpointConfiguration endpointConfiguration) :
                base(ExecutivaPortClient.GetBindingForEndpoint(endpointConfiguration), ExecutivaPortClient.GetEndpointAddress(endpointConfiguration))
        {
            this.Endpoint.Name = endpointConfiguration.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }

        public ExecutivaPortClient(EndpointConfiguration endpointConfiguration, string remoteAddress) :
                base(ExecutivaPortClient.GetBindingForEndpoint(endpointConfiguration), new System.ServiceModel.EndpointAddress(remoteAddress))
        {
            this.Endpoint.Name = endpointConfiguration.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }

        public ExecutivaPortClient(EndpointConfiguration endpointConfiguration, System.ServiceModel.EndpointAddress remoteAddress) :
                base(ExecutivaPortClient.GetBindingForEndpoint(endpointConfiguration), remoteAddress)
        {
            this.Endpoint.Name = endpointConfiguration.ToString();
            ConfigureEndpoint(this.Endpoint, this.ClientCredentials);
        }

        public ExecutivaPortClient(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) :
                base(binding, remoteAddress)
        {
        }

        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        System.Threading.Tasks.Task<SG3ServiceReference.getLiberacaoResponse> SG3ServiceReference.ExecutivaPort.getLiberacaoAsync(SG3ServiceReference.getLiberacaoRequest request)
        {
            return base.Channel.getLiberacaoAsync(request);
        }

        public System.Threading.Tasks.Task<SG3ServiceReference.getLiberacaoResponse> getLiberacaoAsync(SG3ServiceReference.Liberacao liberacao, int tipo_registro)
        {
            SG3ServiceReference.getLiberacaoRequest inValue = new SG3ServiceReference.getLiberacaoRequest();
            inValue.liberacao = liberacao;
            inValue.tipo_registro = tipo_registro;
            return ((SG3ServiceReference.ExecutivaPort)(this)).getLiberacaoAsync(inValue);
        }

        public async System.Threading.Tasks.Task<List<Liberacao>> getLiberacaoWithCredentials(string userName, string passWord, int tipo_registro)
        {

            this.ClientCredentials.UserName.UserName = userName;
            this.ClientCredentials.UserName.Password = passWord;

            //a consulta padrão do SG3 usa um objeto vazio, acredito que para fazer uma pesquisa global sem filtros
            var liberacao = new Liberacao();

            //objeto de retorno
            //SG3ServiceReference.getLiberacaoResponse response;

            //esta etapa é que seta o header do basic auth de forma correta.
            //fonte: https://stackoverflow.com/questions/39354562/consuming-web-service-with-c-sharp-and-basic-authentication
            // using (OperationContextScope scope = new OperationContextScope(this.InnerChannel))
            // {
            //     var httpRequestProperty = new HttpRequestMessageProperty();
            //     httpRequestProperty.Headers[System.Net.HttpRequestHeader.Authorization] = "Basic " +
            //                  Convert.ToBase64String(Encoding.ASCII.GetBytes(this.ClientCredentials.UserName.UserName + ":" +
            //                  this.ClientCredentials.UserName.Password));
            //     OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = httpRequestProperty;

            //     //executa a chamada
            //     response = await this.getLiberacaoAsync(liberacao, tipo_registro);
            // }
            OperationContextScope scope = new OperationContextScope(this.InnerChannel);

            var httpRequestProperty = new HttpRequestMessageProperty();
            httpRequestProperty.Headers[System.Net.HttpRequestHeader.Authorization] = "Basic " +
                         Convert.ToBase64String(Encoding.ASCII.GetBytes(this.ClientCredentials.UserName.UserName + ":" +
                         this.ClientCredentials.UserName.Password));
            OperationContext.Current.OutgoingMessageProperties[HttpRequestMessageProperty.Name] = httpRequestProperty;

            //executa a chamada
            var response = await this.getLiberacaoAsync(liberacao, tipo_registro);

            //converte o array em uma lista
            return response.@return.OfType<Liberacao>().ToList();

        }

        public virtual System.Threading.Tasks.Task OpenAsync()
        {
            return System.Threading.Tasks.Task.Factory.FromAsync(((System.ServiceModel.ICommunicationObject)(this)).BeginOpen(null, null), new System.Action<System.IAsyncResult>(((System.ServiceModel.ICommunicationObject)(this)).EndOpen));
        }

        public virtual System.Threading.Tasks.Task CloseAsync()
        {
            return System.Threading.Tasks.Task.Factory.FromAsync(((System.ServiceModel.ICommunicationObject)(this)).BeginClose(null, null), new System.Action<System.IAsyncResult>(((System.ServiceModel.ICommunicationObject)(this)).EndClose));
        }

        private static System.ServiceModel.Channels.Binding GetBindingForEndpoint(EndpointConfiguration endpointConfiguration)
        {
            if ((endpointConfiguration == EndpointConfiguration.ExecutivaPort))
            {
                System.ServiceModel.BasicHttpsBinding result = new System.ServiceModel.BasicHttpsBinding();
                result.MaxBufferSize = int.MaxValue;
                result.ReaderQuotas = System.Xml.XmlDictionaryReaderQuotas.Max;
                result.MaxReceivedMessageSize = int.MaxValue;
                result.AllowCookies = true;
                result.Security.Mode = System.ServiceModel.BasicHttpsSecurityMode.Transport;
                result.Security.Transport.ClientCredentialType = System.ServiceModel.HttpClientCredentialType.Basic;
                return result;
            }
            throw new System.InvalidOperationException(string.Format("Não foi possível encontrar o ponto de extremidade com o nome \'{0}\'.", endpointConfiguration));
        }

        private static System.ServiceModel.EndpointAddress GetEndpointAddress(EndpointConfiguration endpointConfiguration)
        {
            if ((endpointConfiguration == EndpointConfiguration.ExecutivaPort))
            {
                return new System.ServiceModel.EndpointAddress("https://sg3.executiva.adm.br/femsa/?r=webservice/soap/index");
            }
            throw new System.InvalidOperationException(string.Format("Não foi possível encontrar o ponto de extremidade com o nome \'{0}\'.", endpointConfiguration));
        }

        private static System.ServiceModel.Channels.Binding GetDefaultBinding()
        {
            return ExecutivaPortClient.GetBindingForEndpoint(EndpointConfiguration.ExecutivaPort);
        }

        private static System.ServiceModel.EndpointAddress GetDefaultEndpointAddress()
        {
            return ExecutivaPortClient.GetEndpointAddress(EndpointConfiguration.ExecutivaPort);
        }

        public enum EndpointConfiguration
        {

            ExecutivaPort,
        }
    }
}
