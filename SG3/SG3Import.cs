using AMSLib;
using AMSLib.ACE;
using AMSLib.Common;
using AMSLib.Manager;
using AMSLib.SQL;
using HzBISCommands;
using HzLibConnection.Data;
using Newtonsoft.Json;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools.SG3
{
    public class SG3Info
    {
        public string Persno { get; set; }
        public string Nome { get; set; }
        public string Cardno { get; set; }
        public SG3Import.RE_STATUS Status { get; set; }
    }

    public class SG3Import
    {
        #region Variables
        /// <summary>
        /// Conexão com o BIS.
        /// </summary>
        public BISManager bisManager { get; set; }
        /// <summary>
        /// Conexão com o banco de dados do BIS.
        /// </summary>
        public HzConexao bisConnection { get; set; }
        /// <summary>
        /// Conexão com o banco de dados de Eventos do BIS.
        /// </summary>
        public HzConexao bisLogConnection { get; set; }
        /// <summary>
        /// Log da aplicação.
        /// </summary>
        public ILogger LogTask { get; set; }
        /// <summary>
        /// Log dos resultados.
        /// </summary>
        public ILogger LogResult { get; set; }
        /// <summary>
        /// Coleçao com os hosts do SG3.
        /// </summary>
        public List<SG3Host> sg3Hosts { get; set; }
        #endregion

        #region Enum
        public enum RE_STATUS
        {
            SUCCESS = 0,
            ERROR = 1,
            DELETE = 2,
            EXIST = 3
        }
        #endregion

        #region Mails
        /// <summary>
        /// Envia e-mail.
        /// </summary>
        /// <param name="subject">Título do e-mail.</param>
        /// <param name="body">Corpo do e-mail.</param>
        private void sendMail(HzConexao connection, string subject, string body)
        {
            try
            {
                string mails = "melina.terencio@kof.com.mx;cristiane.msantos@kof.com.mx;tatiane.lobo@kof.com.mx;ronaldo.moscoso@grupoorion.com.br";
                mails = "ronaldomoscoso@gmail.com";
                HzConexao cBIS = new HzConexao(@"10.122.122.100\BIS_2021", "femsasql", "femsasql", "BISEventLog", "System.Data.SqlClient");
                string sql = String.Format("exec msdb..sp_send_dbmail @profile_name = 'Femsa', @recipients = '{0}', @subject = '{1}', @body = '{2}', @body_format = 'html'",
                    mails, subject, body);
                //karen.costa@kof.com.mx;fernanda.santos@kof.com.mx
                cBIS.executeProcedure(sql);
                this.LogTask.Information("E-mail enviado com sucesso.");
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }

        /// <summary>
        /// Envia e-mail com o início da integracao.
        /// </summary>
        /// <param name="connection">Conexão com o banco de dados.</param>
        /// <param name="data">Data da rotina.</param>
        /// <param name="results">Coleção com os resultados da rotina de integracao.</param>
        private void sendStartedMail(HzConexao connection)
        {
            try
            {
                this.LogTask.Information("Enviando e-mail do início da importacao.");

                string subject = String.Format("Importacao SG3 - Rotina iniciada na data {0}", DateTime.Now.ToString(new CultureInfo("pt-BR", true).DateTimeFormat.FullDateTimePattern));
                string HTML = "<html><body>Rotina Iniciada com Sucesso!</body></html>";
                this.sendMail(connection, subject, HTML);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }

        /// <summary>
        /// Envia e-mail com o início da integracao.
        /// </summary>
        /// <param name="connection">Conexão com o banco de dados.</param>
        /// <param name="data">Data da rotina.</param>
        /// <param name="results">Coleção com os resultados da rotina de integracao.</param>
        private void sendErrorMail(HzConexao connection, string errormessage)
        {
            try
            {
                this.LogTask.Information("Enviando e-mail de erro da importacao.");

                string subject = String.Format("Importacao SG3 - Rotina finalizada com erro na data {0}", DateTime.Now.ToString(new CultureInfo("pt-BR", true).DateTimeFormat.FullDateTimePattern));
                string HTML = String.Format("<html><body>Erro: {0}</body></html>", errormessage);
                this.sendMail(connection, subject, HTML);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }

        /// <summary>
        /// Envia e-mail com o início da integracao.
        /// </summary>
        /// <param name="connection">Conexão com o banco de dados.</param>
        /// <param name="results">Coleção com os resultados da rotina de importacao.</param>
        /// <param name="total">Total de registros processados.</param>
        private void sendFinish(HzConexao connection, List<SG3Info> results, int total)
        {
            try
            {
                this.LogTask.Information("Enviando e-mail de termino da importacao.");

                string subject = String.Format("Importacao SG3 - Rotina finalizada na data {0}", DateTime.Now.ToString(new CultureInfo("pt-BR", true).DateTimeFormat.FullDateTimePattern));
                string body = results.Count > 0 ? String.Format("<h2>Total de Registros Processados: {0}</h2><br><h2>Novos Colaboradores: {1}</h2><br><h2>Colaboradores Excluídos: {2}</h2><br><table border=\"1\">", total.ToString(),
                    (results.FindAll(s => s.Status == RE_STATUS.SUCCESS).Count > 0 ? results.FindAll(s => s.Status == RE_STATUS.SUCCESS).Count.ToString() : "0"),
                    (results.FindAll(s => s.Status == RE_STATUS.DELETE).Count > 0 ? results.FindAll(s => s.Status == RE_STATUS.DELETE).Count.ToString() : "0")) :
                    String.Format("<h2>Total de Registros Processados: {0}</h2><br><h2>Não Há Novos Colaboradores.</h2>", total.ToString());
                try
                {
                    if (results.FindAll(s => s.Status == RE_STATUS.SUCCESS).Count > 0)
                    {
                        body += "<tr colspan=\"4\">Novos Colaboradores</tr>";
                        foreach (SG3Info r in results.FindAll(s => s.Status == RE_STATUS.SUCCESS))
                        {
                            body += String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>", r.Persno, r.Nome, r.Cardno, r.Status == RE_STATUS.SUCCESS ? "Ok" : "Erro");
                        }
                    }

                    if (results.FindAll(s => s.Status == RE_STATUS.DELETE).Count > 0)
                    {
                        body += "<tr colspan=\"4\">Colaboradores Excluídos</tr>";
                        foreach (SG3Info r in results.FindAll(s => s.Status == RE_STATUS.DELETE))
                        {
                            body += String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>", r.Persno, r.Nome, r.Cardno, r.Status == RE_STATUS.DELETE ? "Ok" : "Erro");
                        }
                    }

                    if (results.Count > 0)
                        body += "</table>";
                }
                catch
                {

                }
                string HTML = String.Format("<html><body>{0}</body></html>", body);
                this.sendMail(connection, subject, HTML);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }
        #endregion

        #region Functions
        /// <summary>
        /// Retorna o clientid da divisão no BIS relativo
        /// ao código de área do SAP.
        /// </summary>
        /// <param name="area">Descricao da área.</param>
        /// <returns></returns>
        private string getClientidFromArea(string area)
        {
            string retval = null;
            try
            {
                this.LogTask.Information(String.Format("Verificando o clientid da area: {0}", area.TrimStart(new Char[] { '0' })));
                using (DataTable table = this.bisConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where area = '{0}'", area.TrimStart(new Char[] { '0' }))))
                {
                    if (table.Rows.Count > 0)
                    {
                        retval = table.Rows[0]["clientid"].ToString();
                        this.LogTask.Information(String.Format("Area encontrada: {0}", retval));
                    }
                    else
                        this.LogTask.Information("Area nao encontrada.");
                }

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return null;
            }
        }

        public async void Import()
        {
            BSPersons persons = null;
            List<SG3Info> sg3Info = new List<SG3Info>();
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\ImportSG3TaskLog.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.LogResult = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\ImportSG3ResultLog.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.sendStartedMail(this.bisConnection);
                int total = 0;
                persons = new BSPersons(this.bisManager, this.bisConnection);
                foreach (SG3Host host in this.sg3Hosts)
                {
                    User user = new User() { username = host.Username, password = host.Password };
                    using (var client = new HttpClient())
                    {
                        this.LogTask.Information(String.Format("Pesquisando a base: {0}", host.Username));
                        this.LogTask.Information(String.Format("{0}/{1}", host.Host, host.LoginCommand), new StringContent(JsonConvert.SerializeObject(user), Encoding.UTF8, "application/json"));
                        HttpResponseMessage message = await client.PostAsync(String.Format("{0}/{1}", host.Host, host.LoginCommand), new StringContent(JsonConvert.SerializeObject(user), Encoding.UTF8, "application/json"));
                        if (message.StatusCode != System.Net.HttpStatusCode.OK)
                        {
                            this.LogTask.Information(String.Format("Message: {0}", message.StatusCode.ToString()));
                            continue;
                        }
                        string response = await message.Content.ReadAsStringAsync();
                        if (String.IsNullOrEmpty(response))
                        {
                            this.LogTask.Information("Nenuma informacao encontrada.");
                            continue;
                        }

                        user = (User)JsonConvert.DeserializeObject<User>(response);

                        var httpClient = new HttpClient();

                        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", user.token);

                        // Send the request and get the response
                        httpClient.Timeout = TimeSpan.FromMinutes(10);
                        var rr = await httpClient.GetAsync(new Uri(String.Format("{0}/{1}", host.Host, host.Command)));

                        // Handle the response
                        var rc = rr.Content.ReadAsStringAsync();
                        var liberacao = JsonConvert.DeserializeObject<List<Liberacao>>(rc.Result);
                        this.LogTask.Information(String.Format("{0} registros encontrados.", liberacao.Count.ToString()));
                        int count = 1;
                        total += liberacao.Count;
                        foreach (Liberacao l in liberacao)
                        {
                            this.LogTask.Information(String.Format("{0} de {1}, RG: {2}, CPF: {3}, Nome: {4}", (count++).ToString(), liberacao.Count.ToString(), l.rg, l.cpf, l.colaborador.Replace("'", "")));

                            if (l.tipo_terceiro == null)
                                continue;

                            if (l.tipo_terceiro.ToLower().Equals("aprendiz") || l.tipo_terceiro.ToLower().Equals("fixo") || l.tipo_terceiro.ToLower().Equals("volante") ||
                            l.tipo_terceiro.ToLower().Equals("outros contratos fixo") || l.tipo_terceiro.ToLower().Equals("socio fixo") || l.tipo_terceiro.ToLower().Equals("cooperfemsa") ||
                            l.tipo_terceiro.ToLower().Equals("estagiario") || l.tipo_terceiro.ToLower().Equals("outros contratos aprendiz") || l.tipo_terceiro.ToLower().Equals("residente") ||
                            l.tipo_terceiro.ToLower().Equals("outros contratos") || l.tipo_terceiro.ToLower().Equals("temporario 6019") ||
                            l.tipo_terceiro.ToLower().Equals("socio residente"))
                            {
                                if (host.Username.ToLower().Equals("webservice.femsabrasil") && l.liberado.Equals("S"))
                                {
                                    this.LogTask.Information("Importando....");
                                    sg3Info.Add(this.Import(l));
                                }

                                if (host.Username.ToLower().Equals("webservice.femsa") || host.Username.ToLower().Equals("webservice.femsat1"))
                                {
                                    if (String.IsNullOrEmpty(l.cpf))
                                    {
                                        this.LogTask.Information("CPF em branco.");
                                        continue;
                                    }
                                    if (l.tipo_terceiro.ToLower().Equals("residente") || l.tipo_terceiro.ToLower().Equals("socio residente"))
                                    {
                                        this.LogTask.Information("Importando....");
                                        sg3Info.Add(this.Import(l));
                                    }
                                    else
                                    {
                                        this.LogTask.Information("Excluindo....");
                                        sg3Info.Add(this.Delete(l));
                                    }
                                }
                            }
                        }
                    }
                    //var x = liberacao.FindAll(delegate (Liberacao l) {
                    //        return l.liberado.Equals("S") && l.status_alocacao != null && l.status_alocacao.Equals("S") &&
                    //        (l.tipo_terceiro.ToLower().Equals("aprendiz") || l.tipo_terceiro.ToLower().Equals("fixo") || l.tipo_terceiro.ToLower().Equals("volante") ||
                    //        l.tipo_terceiro.ToLower().Equals("outros contratos fixo") || l.tipo_terceiro.ToLower().Equals("socio fixo") || l.tipo_terceiro.ToLower().Equals("cooperfemsa") ||
                    //        l.tipo_terceiro.ToLower().Equals("estagiario") || l.tipo_terceiro.ToLower().Equals("outros contratos aprendiz")); });
                    //    import.Import(x);
                    //}
                }
                this.sendFinish(this.bisConnection, sg3Info, total);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                this.sendErrorMail(this.bisConnection, ex.Message);
            }
            finally
            {
                this.LogTask.Information("Descontecando o BIS.");
                if (this.bisManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }
        /// <summary>
        /// Importa os dados do SG3.
        /// </summary>
        /// <param name="liberacao">Classe com os dados do SG3.</param>
        public SG3Info Import(Liberacao liberacao)
        {
            SG3Info retval = null;
            try
            {
                string firstname = "";
                string lastname = "";
                BSCommon.ParseName(liberacao.colaborador.Replace("'", ""), out firstname, out lastname);

                string sql = String.Format("select persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), persclassid from bsuser.persons where status = 1 and persclass = 'E' and persno = '{0}'",
                    liberacao.cpf.Replace("'", ""));
                using (DataTable table = this.bisConnection.loadDataTable(sql))
                {
                    if (table.Rows.Count > 0)
                    {
                        this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                        retval = new SG3Info() { Persno = liberacao.cpf, Nome = liberacao.colaborador, Status = RE_STATUS.EXIST };
                    }
                    else
                    {
                        this.LogTask.Information("Nova pessoa.");

                        //string clientid = this.getClientidFromArea(liberacao.estabelecimento);
                        //if (!String.IsNullOrEmpty(clientid))
                        //{
                        //    this.LogTask.Information(String.Format("Alterando a divisao para o ID: {0}", clientid));
                        //    if (this.bisManager.SetClientID(clientid))
                        //        this.LogTask.Information("ClientID alterado com sucesso.");
                        //    else
                        //        this.LogTask.Information(String.Format("Erro ao alterar o ClientID. {0}", this.bisManager.GetErrorMessage()));
                        //}
                        using (BSPersons persons = new BSPersons(this.bisManager, this.bisConnection))
                        {
                            persons.PERSNO = liberacao.cpf.Replace("'", "").Replace(".", "").Replace("-", "");
                            persons.FIRSTNAME = firstname;
                            persons.LASTNAME = lastname;
                            persons.GRADE = "SG3";
                            persons.PERSCLASSID = "0013B4437A2455E1";
                            persons.JOB = liberacao.funcao.Replace("'", "");
                            persons.COMPANYID = BSSQLCompanies.GetIDCNPJ(this.bisConnection, liberacao.cnpj);

                            if (persons.Save() == BS_SAVE.SUCCESS)
                            {
                                this.LogTask.Information("Pessoa gravada com sucesso.");
                                retval = new SG3Info() { Persno = persons.PERSNO, Nome = liberacao.colaborador, Status = RE_STATUS.SUCCESS };
                            }
                            else
                            {
                                this.LogTask.Information("Erro ao gravar a pessoa.");
                                retval = new SG3Info() { Persno = persons.PERSNO, Nome = liberacao.colaborador, Status = RE_STATUS.ERROR };
                            }
                        }
                    }
                }

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return new SG3Info() { Persno = liberacao.cpf, Nome = liberacao.colaborador, Status = RE_STATUS.ERROR };
            }
        }

        /// <summary>
        /// Deleta os dados do SG3.
        /// </summary>
        /// <param name="liberacao">Classe com os dados do SG3.</param>
        public SG3Info Delete(Liberacao liberacao)
        {
            SG3Info retval = null;
            try
            {
                string sql = String.Format("select persid from bsuser.persons where status = 1 and persclass = 'E' and persno = '{0}' and grade = 'SG3' and persclassid = '0013B4437A2455E1'",
                    liberacao.cpf.Replace("'", ""));
                using (DataTable table = this.bisConnection.loadDataTable(sql))
                {
                    if (table.Rows.Count > 0)
                    {
                        this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                        using (BSPersons persons = new BSPersons(this.bisManager, this.bisConnection))
                        {
                            this.LogTask.Information("Excluindo a pessoa.");
                            if (persons.DeletePerson(table.Rows[0]["persid"].ToString()) == BS_ADD.SUCCESS)
                            {
                                this.LogTask.Information("Pessoa excluida com sucesso.");
                                retval = new SG3Info() { Persno = persons.PERSNO, Nome = liberacao.colaborador, Status = RE_STATUS.DELETE };
                            }
                            else
                            {
                                this.LogTask.Information("Erro ao excluir a pessoa.");
                                retval = new SG3Info() { Persno = persons.PERSNO, Nome = liberacao.colaborador, Status = RE_STATUS.ERROR };
                            }
                        }
                    }
                }

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return new SG3Info() { Persno = liberacao.cpf, Nome = liberacao.colaborador, Status = RE_STATUS.ERROR };
            }
        }
        #endregion

        #region Events
        /// <summary>
        /// Construtor da classe.
        /// </summary>
        /// <param name="logtask">Log dos eventos.</param>
        /// <param name="logresult">Log dos resultados.</param>
        /// <param name="bissql">Parâmetros do SQL Server do BIS.</param>
        /// <param name="bislogsql">Parâmetros do SQL Server de Eventos do BIS.</param>
        /// <param name="bismanager">Objeto do BIS.</param>
        /// <param name="hosts">Coleçao dos hosts SG3.</param>
        public SG3Import(ILogger logtask, ILogger logresult, SQLParameters bissql, SQLParameters bislogsql, BISManager bismanager, List<SG3Host> hosts)
        {
            this.LogTask = logtask;
            this.LogResult = logresult;
            this.bisConnection = new HzConexao(bissql.SQLHost, bissql.SQLUser, bissql.SQLPwd, "acedb", "System.Data.SqlClient");
            this.bisLogConnection = new HzConexao(bislogsql.SQLHost, bislogsql.SQLUser, bislogsql.SQLPwd, "BISEventLog", "System.Data.SqlClient");
            this.bisManager = bismanager;
            this.sg3Hosts = hosts;
        }
        #endregion  
    }
}
