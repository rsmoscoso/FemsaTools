using AMSLib;
using AMSLib.ACE;
using AMSLib.Common;
using AMSLib.Manager;
using AMSLib.SQL;
using HzBISCommands;
using HzLibConnection.Data;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools
{
    public class FemsaInfo
    {
        public string Persno { get; set; }
        public string Nome { get; set; }
        public string Cardno { get; set; }
        public FEMSAImport.RE_STATUS Status { get; set; }
    }
    public class FEMSAImport
    {
        private enum IMPORT
        {
            IMPORT_SAMERE = 0,
            IMPORT_DIFFRE = 1,
            IMPORT_NEW = 2,
            IMPORT_NOPHOTO = 3,
            IMPORT_UNDELETE = 4,
            IMPORT_DISABLED = 5,
            IMPORT_DOUBLERE = 6,
            IMPORT_NONE = 7,
            IMPORT_DELETE = 8
        }

        public enum RE_STATUS
        {
            SUCCESS = 0,
            ERROR = 1,
            DELETE = 2
        }
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
        #endregion

        #region Functions
        /// <summary>
        /// Envia e-mail.
        /// </summary>
        /// <param name="subject">Título do e-mail.</param>
        /// <param name="body">Corpo do e-mail.</param>
        private void sendMail(HzConexao connection, string subject, string body)
        {
            try
            {
                HzConexao cBIS = new HzConexao(@"10.122.122.100\BIS_2021", "femsasql", "femsasql", "BISEventLog", "System.Data.SqlClient");
                string sql = String.Format("exec msdb..sp_send_dbmail @profile_name = 'Femsa', @recipients = 'ronaldo.moscoso@grupoorion.com.br', @subject = '{0}', @body = '{1}', @body_format = 'html'",
                    subject, body);
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
        private void sendStartedMail(HzConexao connection, string text)
        {
            try
            {
                this.LogTask.Information("Enviando e-mail do início da integracao.");

                string subject = text;
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
                this.LogTask.Information("Enviando e-mail de erro da integracao.");

                string subject = String.Format("Importacao SAP - Rotina finalizada com erro na data {0}", DateTime.Now.ToString(new CultureInfo("pt-BR", true).DateTimeFormat.FullDateTimePattern));
                string HTML = String.Format("<html><body>Erro: {0}</body></html>", errormessage);
                this.sendMail(connection, subject, HTML);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }

        /// <summary>
        /// Apaga a pessoa e seus cartões.
        /// </summary>
        /// <param name="persid">ID da pessoa.</param>
        /// <returns></returns>
        private bool deletePerson(string persid)
        {
            try
            {
                BSPersons per = new BSPersons(this.bisManager, this.bisConnection);
                this.LogTask.Information("Excluindo o colaborador...");
                if (per.DeletePerson(persid) == BS_ADD.SUCCESS)
                    this.LogTask.Information("Colaborador excluido com sucesso.");
                else
                    this.LogTask.Information(String.Format("Erro ao excluir o colaborador: {0}", per.GetErrorMessage()));

                return true;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Envia e-mail com o início da integracao.
        /// </summary>
        /// <param name="connection">Conexão com o banco de dados.</param>
        /// <param name="results">Coleção com os resultados da rotina de integracao.</param>
        /// <param name="total">Total de registros processados.</param>
        private void sendFinish(HzConexao connection, List<FemsaInfo> results, int total)
        {
            try
            {
                this.LogTask.Information("Enviando e-mail de termino da integracao.");

                string subject = String.Format("Importacao SAP - Rotina finalizada na data {0}", DateTime.Now.ToString(new CultureInfo("pt-BR", true).DateTimeFormat.FullDateTimePattern));
                string body = results.Count > 0 ? String.Format("<h2>Total de Registros Processados: {0}</h2><br><h2>Novos Colaboradores: {1}</h2><br><h2>Colaboradores Excluídos: {2}</h2><br><table border=\"1\">", total.ToString(),
                    (results.FindAll(s => s.Status == RE_STATUS.SUCCESS).Count > 0 ? results.FindAll(s => s.Status == RE_STATUS.SUCCESS).Count.ToString() : "0"),
                    (results.FindAll(s => s.Status == RE_STATUS.DELETE).Count > 0 ? results.FindAll(s => s.Status == RE_STATUS.DELETE).Count.ToString() : "0")) :
                    String.Format("<h2>Total de Registros Processados: {0}</h2><br><h2>Não Há Novos Colaboradores.</h2>", total.ToString());
                try
                {
                    if (results.FindAll(s => s.Status == RE_STATUS.SUCCESS).Count > 0)
                    {
                        body += "<tr colspan=\"4\">Novos Colaboradores</tr>";
                        foreach (FemsaInfo r in results.FindAll(s => s.Status == RE_STATUS.SUCCESS))
                        {
                            body += String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>", r.Persno, r.Nome, !String.IsNullOrEmpty(r.Cardno) ? r.Cardno : "Sem Cartao", r.Status == RE_STATUS.SUCCESS ? "Ok" : "Erro");
                        }
                    }

                    if (results.FindAll(s => s.Status == RE_STATUS.DELETE).Count > 0)
                    {
                        body += "<tr colspan=\"4\">Colaboradores Excluídos</tr>";
                        foreach (FemsaInfo r in results.FindAll(s => s.Status == RE_STATUS.DELETE))
                        {
                            body += String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>", r.Persno, r.Nome, !String.IsNullOrEmpty(r.Cardno) ? r.Cardno : "Sem Cartao", r.Status == RE_STATUS.DELETE ? "Ok" : "Erro");
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

        /// <summary>
        /// Compara o PERSNO do ID desejado com outro valor de PERSNO.
        /// </summary>
        /// <param name="persid">ID da pessoa.</param>
        /// <param name="persno">Persno a ser comparado.</param>
        /// <returns></returns>
        private bool? doubleCheckRE(string persid, string persno)
        {
            try
            {
                string persnocheck = BSSQLPersons.GetPersno(this.bisConnection, persno);
                return !String.IsNullOrEmpty(persnocheck).Equals(persno);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return null;
            }
        }

        /// <summary>
        /// Verifica a existência do RE.
        /// </summary>
        /// <param name="re">RE do colaborador.</param>
        /// <returns></returns>
        private IMPORT existRE(string persno, string nome, string status, out BSEventsInfo events)
        {
            IMPORT retval = IMPORT.IMPORT_NONE;
            events = new BSEventsInfo();
            try
            {
                if (String.IsNullOrEmpty(status))
                    status = "";

                this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Status SAP: {2}", nome, persno, status));

                using (DataTable tbl = BSSQLPersons.GetID(this.bisConnection, persno, BS_STATUS.ACTIVE))
                {
                    if (tbl != null && tbl.Rows.Count == 1)
                    {
                        if (tbl.Rows[0]["status"].ToString().Equals("0") && status.ToLower().Equals("saiu da empresa"))
                        {
                            retval = IMPORT.IMPORT_DISABLED;
                            this.LogTask.Information("Pessoa excluida.");
                        }
                        else if (tbl.Rows[0]["status"].ToString().Equals("1") && status.ToLower().Equals("saiu da empresa"))
                        {
                            retval = IMPORT.IMPORT_DELETE;
                            this.deletePerson(tbl.Rows[0]["persid"].ToString());
                        }
                        else
                        {
                            events.PERSID = tbl.Rows[0]["persid"].ToString();
                            retval = IMPORT.IMPORT_SAMERE;
                            this.LogTask.Information("Pessoa existente.");
                        }
                    }
                    else if (tbl.Rows.Count > 0 && (status.ToLower().Equals("ativo") || status.ToLower().Equals("s")))
                    {
                        this.LogTask.Information("Pessoa com persno repetido. Ultimos cardid, persid e data de acesso ativos.");
                        events = BSSQLEvents.GetLastCardidFromPersno(this.bisLogConnection, persno);

                        if (events == null || String.IsNullOrEmpty(events.CARDID))
                            this.LogTask.Information("Nao ha eventos para essa pessoa.");
                        else
                        {
                            events.PERSID = BSSQLCards.GetPersID(this.bisConnection, events.CARDID);
                            this.LogTask.Information(String.Format("Pessoa com persno repetido. Ultimos cardid, persid e data de acesso ativos: {0}, {1}", events.CARDID, events.PERSID, events.DATAACESSO));
                            if (this.doubleCheckRE(events.PERSID, persno).Value)
                            {
                                retval = IMPORT.IMPORT_DOUBLERE;
                                this.LogTask.Information("Double check ok.");
                            }
                            else
                            {
                                retval = IMPORT.IMPORT_DIFFRE;
                                this.LogTask.Information("Erro no double check. Registro sera ignorado.");
                            }
                        }
                    }
                    else if (tbl.Rows.Count > 0 && status.ToLower().Equals("saiu da empresa"))
                    {
                        foreach(DataRow row in tbl.Rows)
                        {
                            if (row["status"].ToString().Equals("1"))
                            {
                                this.deletePerson(row["persid"].ToString());
                                retval = IMPORT.IMPORT_DELETE;
                            }
                        }
                    }
                    else if (tbl.Rows.Count == 0 && status.ToLower().Equals("saiu da empresa"))
                    {
                        retval = IMPORT.IMPORT_DISABLED;
                        this.LogTask.Information("Pessoa excluida.");
                    }
                    else
                    {
                        retval = IMPORT.IMPORT_NEW;
                        this.LogTask.Information("Pessoa inexistente. Seguindo a importacao.");
                    }
                }

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return retval;
            }
        }

        /// <summary>
        /// Atualiza o idnumber para o último cardid utilizado 
        /// para registros duplicados.
        /// </summary>
        /// <param name="persid">ID da pessoa.</param>
        /// <param name="cardid">Id do cartão.</param>
        /// <returns></returns>
        private bool updateCardidDouble(string persid, string cardid)
        {
            bool retval = false;
            try
            {
                this.LogTask.Information(String.Format("Alterando o idnumber para o ultimo cardid utilziado. CardID: {0}", cardid));
                string sql = String.Format("update bsuser.persons set idnumber = '{0}' where persid = '{1}'",
                    cardid, persid);
                if (this.bisConnection.executeProcedure(sql))
                    this.LogTask.Information("Registro gravado com sucesso.");
                else
                    this.LogTask.Information("Erro ao gravar o registro.");
                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return retval;
            }
        }
        /// <summary>
        /// Atualiza os dados do SQL sem necessidade de broadcast
        /// para as controladoras.
        /// </summary>
        /// <param name="personsinfo">Classe com as informações da pessoa.</param>
        /// <returns></returns>
        private bool updateOnlySQL(BSPersonsInfoFEMSA personsinfo)
        {
            bool retval = false;
            try
            {
                this.LogTask.Information("Alterando o registro existente diretamente no ACEDB.");
                string sql = String.Format("update bsuser.persons set departmattend = '{0}', department = '{1}', costcentre = '{2}', centraloffice = '{3}', job = '{4}', cityofbirth = '{5}', attendant = '{6}' where persid = '{7}'",
                    Global.getString(personsinfo.DEPARTMATTEND, 40), Global.getString(personsinfo.DEPARTMENT, 40), personsinfo.ACTIVITY, Global.getString(personsinfo.CENTRALOFFICE, 40), 
                    Global.getString(personsinfo.JOB, 40), Global.getString(personsinfo.COSTCENTRE, 32), Global.getString(personsinfo.DESCRICAOAREA, 40), personsinfo.PERSID);
                if (this.bisConnection.executeProcedure(sql))
                    this.LogTask.Information("Registro gravado com sucesso.");
                else
                    this.LogTask.Information("Erro ao gravar o registro.");

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return retval;
            }
        }

        /// <summary>
        /// Carrega o objeto ACE do BIS com as suas propriedades.
        /// </summary>
        /// <param name="bisoject">Objeto BIS a ser carregado com as propriedades: BSPersons, BSVisitors, BSCards, etc...</param>
        /// <param name="properties">Propriedades referentes ao objecto BIS.</param>
        /// <returns>Retorna o objeto BIS com as suas propriedades preenchidas. Se houver erro, retorna null.</returns>
        private object FillBISObject(object bisoject, object properties)
        {
            try
            {
                foreach (FieldInfo p in BSCommon.GetFields(bisoject))
                {
                    foreach (PropertyInfo pi in ((BSPersonsInfoFEMSA)properties).GetType().GetProperties())
                    {
                        if (p.Name.Equals(pi.Name))
                        {
                            var value = pi.GetValue(properties);
                            if (value == null)
                                continue;
                            if (p.FieldType.Equals(typeof(ACEDateT)) && pi.GetType().Equals(p.FieldType))
                                p.SetValue(bisoject, value ?? new ACEDateT());
                            else if (!pi.GetType().Equals(p.FieldType) && p.FieldType.Equals(typeof(ACEDateT)))
                            {
                                if (value != null)
                                {
                                    ACEDateT dt = new ACEDateT((uint)((DateTime)value).Day, (uint)((DateTime)value).Month, (uint)((DateTime)value).Year);
                                    if (dt.m_iYear > 1)
                                        p.SetValue(bisoject, dt);
                                }
                            }
                            else if (p.FieldType.Equals(typeof(DateTime)) && pi.PropertyType.Equals(typeof(System.String)))
                            {
                                if (value != null)
                                {
                                    value = DateTime.Parse(value.ToString());
                                    p.SetValue(bisoject, value);
                                }

                            }
                            else if (p.FieldType.Equals(typeof(ACEMaritalStateT)))
                            {
                                p.SetValue(bisoject, (int)value);
                            }
                            else
                                p.SetValue(bisoject, value);
                            break;
                        }
                    }
                }

                return bisoject;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Retorna o clientid da divisão no BIS relativo
        /// ao código de área do SAP.
        /// </summary>
        /// <param name="codigo">Código da área.</param>
        /// <returns></returns>
        private string getClientidFromArea(string codigo)
        {
            string retval = null;
            try
            {
                this.LogTask.Information(String.Format("Verificando o clientid da area: {0}", codigo.TrimStart(new Char[] { '0' })));
                using (DataTable table = this.bisConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where codigo = '{0}'", codigo.TrimStart(new Char[] { '0' }))))
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

        /// <summary>
        /// Retorna o clientid da divisão no BIS relativo
        /// ao código de área do SAP.
        /// </summary>
        /// <param name="codigo">Código da área.</param>
        /// <returns></returns>
        private string getClientid(string area)
        {
            string retval = null;
            try
            {
                this.LogTask.Information(String.Format("Verificando o clientid da area: {0}", area));
                using (DataTable table = this.bisConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where area = '{0}'", area)))
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
        /// <summary>
        /// Salva o registro no BIS.
        /// </summary>
        /// <param name="personsinfo">Classe com os dados da pessoa.</param>
        /// <returns></returns>
        private RE_STATUS save(BSPersonsInfoFEMSA personsinfo)
        {
            RE_STATUS retval = RE_STATUS.ERROR;
            try
            {
                BSPersons persons = new BSPersons(this.bisManager, this.bisConnection);
                string firstname = null;
                string lastname = null;
                
                BSCommon.ParseName(personsinfo.NOME, out firstname, out lastname);
                personsinfo.FIRSTNAME = firstname;
                personsinfo.LASTNAME = lastname;

                this.LogTask.Information("Carregando o objeto BSPersons.");
                persons = (BSPersons)this.FillBISObject(persons, personsinfo);
                if (!String.IsNullOrEmpty(personsinfo.PERSID))
                {
                    this.LogTask.Information("Atualizando a pessoa.");
                    this.LogTask.Information("Carregando a pessoa.");
                    if (persons.Load(personsinfo.PERSID))
                        this.LogTask.Information("Pessoa carregada com sucesso.");
                    else
                        this.LogTask.Information("Erro ao carregar a pessoa.");
                }
                else
                    this.LogTask.Information("Nova pessoa.");

                string clientid = this.getClientidFromArea(personsinfo.CODIGOAREA);
                if (!String.IsNullOrEmpty(clientid))
                {
                    this.LogTask.Information(String.Format("Alterando a divisao para o ID: {0}", clientid));
                    if (this.bisManager.SetClientID(clientid))
                        this.LogTask.Information("ClientID alterado com sucesso.");
                    else
                        this.LogTask.Information(String.Format("Erro ao alterar o ClientID. {0}", this.bisManager.GetErrorMessage()));
                }

                if ((String.IsNullOrEmpty(personsinfo.PERSID) || (!String.IsNullOrEmpty(persons.GetPersonId())) ? persons.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS : false))
                {
                    this.LogTask.Information(String.Format("Registro gravado com sucesso. PERSID: {0}", persons.GetPersonId()));
                    personsinfo.PERSID = persons.GetPersonId();
                }
                else
                    throw new Exception(String.Format("Erro ao gravar o registro. {0}", persons.GetErrorMessage()));

                if (!String.IsNullOrEmpty(personsinfo.CARDNO))
                {
                    this.LogTask.Information(String.Format("Atribuindo o cartao {0}", personsinfo.CARDNO));
                    this.LogTask.Information("Verificando a existencia do cartao.");
                    string cardid = BSSQLCards.GetID(this.bisConnection, personsinfo.CARDNO, "1568", BS_STATUS.ACTIVE);
                    if (String.IsNullOrEmpty(cardid))
                    {
                        this.LogTask.Information("Criando o novo cartao.");
                        BSCards cards = new BSCards(this.bisManager);
                        cards.CODEDATA = "1568";
                        cards.CARDNO = personsinfo.CARDNO;
                        if (cards.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                        {
                            this.LogTask.Information("Cartao criado com sucesso.");
                            this.LogTask.Information("Verificando o ID do cartao.");
                            cardid = BSSQLCards.GetID(this.bisConnection, personsinfo.CARDNO, "1568", BS_STATUS.ACTIVE);
                            this.LogTask.Information("ID encontrado: {0}. Atribuindo o cartao.", cardid);
                            if (persons.AttachCard(cardid) == BS_ADD.SUCCESS)
                                this.LogTask.Information("Cartao atribuido com sucesso.");
                            else
                                throw new Exception("Erro ao atribuir o cartao.");
                        }
                        else
                            throw new Exception(String.Format("Erro ao criar o cartao. {0}", cards.GetErrorMessage()));
                        retval = RE_STATUS.SUCCESS;
                    }
                    else
                    {
                        retval = RE_STATUS.ERROR;
                        string persid = BSSQLCards.GetPersID(this.bisConnection, cardid);
                        string persno = BSSQLPersons.GetPersno(this.bisConnection, persid);
                        this.LogTask.Information(String.Format("Cartao atribuido a outro colaborador. PERSID: {0}, PERSNO: {1}", persid, persno));
                    }
                }

                this.updateCustomField(persons, personsinfo);

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return retval;
            }
        }

        /// <summary>
        /// Salva o registro no BIS.
        /// </summary>
        /// <param name="personsinfo">Classe com os dados da pessoa.</param>
        /// <returns></returns>
        private RE_STATUS save(BSPersonsInfoSG3 personsinfo)
        {
            RE_STATUS retval = RE_STATUS.ERROR;
            try
            {
                BSPersons persons = new BSPersons(this.bisManager, this.bisConnection);
                string firstname = null;
                string lastname = null;

                BSCommon.ParseName(personsinfo.NOME, out firstname, out lastname);
                personsinfo.FIRSTNAME = firstname;
                personsinfo.LASTNAME = lastname;

                this.LogTask.Information("Carregando o objeto BSPersons.");
                persons = (BSPersons)this.FillBISObject(persons, personsinfo);
                if (!String.IsNullOrEmpty(personsinfo.PERSID))
                {
                    this.LogTask.Information("Atualizando a pessoa.");
                    this.LogTask.Information("Carregando a pessoa.");
                    if (persons.Load(personsinfo.PERSID))
                        this.LogTask.Information("Pessoa carregada com sucesso.");
                    else
                        this.LogTask.Information("Erro ao carregar a pessoa.");
                }
                else
                    this.LogTask.Information("Nova pessoa.");

                if (!String.IsNullOrEmpty(personsinfo.COMPANYNO))
                {
                    this.LogTask.Information(String.Format("Verificando o ID da empresa.", personsinfo.COMPANYNO));
                    string companyid = BSSQLCompanies.GetID(this.bisConnection, personsinfo.COMPANYNO);
                    if (!String.IsNullOrEmpty(companyid))
                    {
                        this.LogTask.Information(String.Format("Empresa encontrada. ID: {0}", companyid));
                        personsinfo.COMPANYID = companyid;
                    }
                    else
                        this.LogTask.Information("Empresa nao encontrada.");
                }

                string clientid = this.getClientid(personsinfo.ESTABELECIMENTO);
                if (!String.IsNullOrEmpty(clientid))
                {
                    this.LogTask.Information(String.Format("Alterando a divisao para o ID: {0}", clientid));
                    if (this.bisManager.SetClientID(clientid))
                        this.LogTask.Information("ClientID alterado com sucesso.");
                    else
                        this.LogTask.Information(String.Format("Erro ao alterar o ClientID. {0}", this.bisManager.GetErrorMessage()));
                }

                if ((String.IsNullOrEmpty(personsinfo.PERSID) || (!String.IsNullOrEmpty(persons.GetPersonId())) ? persons.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS : false))
                {
                    this.LogTask.Information(String.Format("Registro gravado com sucesso. PERSID: {0}", persons.GetPersonId()));
                    personsinfo.PERSID = persons.GetPersonId();
                }
                else
                    throw new Exception(String.Format("Erro ao gravar o registro. {0}", persons.GetErrorMessage()));

                if (!String.IsNullOrEmpty(personsinfo.CARDNO))
                {
                    this.LogTask.Information(String.Format("Atribuindo o cartao {0}", personsinfo.CARDNO));
                    this.LogTask.Information("Verificando a existencia do cartao.");
                    string cardid = BSSQLCards.GetID(this.bisConnection, personsinfo.CARDNO, "1568", BS_STATUS.ACTIVE);
                    if (String.IsNullOrEmpty(cardid))
                    {
                        this.LogTask.Information("Criando o novo cartao.");
                        BSCards cards = new BSCards(this.bisManager);
                        cards.CODEDATA = "1568";
                        cards.CARDNO = personsinfo.CARDNO;
                        if (cards.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                        {
                            this.LogTask.Information("Cartao criado com sucesso.");
                            this.LogTask.Information("Verificando o ID do cartao.");
                            cardid = BSSQLCards.GetID(this.bisConnection, personsinfo.CARDNO, "1568", BS_STATUS.ACTIVE);
                            this.LogTask.Information("ID encontrado: {0}. Atribuindo o cartao.", cardid);
                            if (persons.AttachCard(cardid) == BS_ADD.SUCCESS)
                                this.LogTask.Information("Cartao atribuido com sucesso.");
                            else
                                throw new Exception("Erro ao atribuir o cartao.");
                        }
                        else
                            throw new Exception(String.Format("Erro ao criar o cartao. {0}", cards.GetErrorMessage()));
                        retval = RE_STATUS.SUCCESS;
                    }
                    else
                    {
                        retval = RE_STATUS.ERROR;
                        string persid = BSSQLCards.GetPersID(this.bisConnection, cardid);
                        string persno = BSSQLPersons.GetPersno(this.bisConnection, persid);
                        this.LogTask.Information(String.Format("Cartao atribuido a outro colaborador. PERSID: {0}, PERSNO: {1}", persid, persno));
                    }
                }

                this.updateCustomField(persons, personsinfo);

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return retval;
            }
        }

        /// <summary>
        /// Carrega a pessoa para os campos adicionais.
        /// </summary>
        /// <param name="persons">Classe da pessoa.</param>
        /// <param name="persid">ID da pessoa.</param>
        /// <returns></returns>
        private BSPersons loadPerson(BSPersons persons, string persid)
        {
            try
            {
                if (persons == null || !persons.GetPersonId().Equals(persid))
                {
                    persons = new BSPersons(this.bisManager, this.bisConnection);
                    this.LogTask.Information("Carregando a pessoa.");
                    if (persons.Load(persid))
                        this.LogTask.Information("Pessoa carregada com sucesso.");
                    else
                        throw new Exception("Erro ao carregar a pessoa.");
                }

                return persons;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return null;
            }
        }
        /// <summary>
        /// Aponta o valor do customfield no BIS.
        /// </summary>
        /// <param name="persons">Classe da pessoa.</param>
        /// <param name="customfield">Nome do campo customizável.</param>
        /// <param name="type">Tipo do campo.</param>
        /// <param name="value">Valor do campo.</param>
        private bool setCustomFieldBIS(BSPersons persons, string persid, string customfield, string type, string value, out BSPersons personsout)
        {
            bool retval = false;
            try
            {
                this.LogTask.Information(String.Format("Alterando o campo {0} para {1}", customfield, value));
                if (String.IsNullOrEmpty(persons.GetPersonId()))
                {
                    if (String.IsNullOrEmpty(((persons = this.loadPerson(persons, persid)).GetPersonId())))
                        throw new Exception(persons.GetErrorMessage());
                }
                if (retval = persons.SetCustomField(customfield, type, value))
                    this.LogTask.Information("Campo alterado com sucesso.");
                else
                    this.LogTask.Information(String.Format("Erro ao setar {0} no BIS.", customfield));

                personsout = persons;

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                throw new Exception("Erro ao carregar a pessoa.");
            }
        }

        /// <summary>
        /// Aponta o valor do customfield no SQL.
        /// </summary>
        /// <param name="persid">ID da pessoa.</param>
        /// <param name="customfield">Nome do campo customizável.</param>
        /// <param name="value">Valor do campo.</param>
        private void setCustomFieldSQL(string persid, string customfield, string value)
        {
            try
            {
                this.LogTask.Information(String.Format("Alterando o campo {0} para {1}", customfield, value));
                if (!BSSQLPersons.SetAdditionalField(this.bisConnection, persid, customfield, value))
                    this.LogTask.Information(String.Format("Erro ao gravar {0} no BIS.", customfield));
                else
                    this.LogTask.Information("Campo alterado com sucesso.");
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Atualiza os campos customizáveis.
        /// </summary>
        /// <param name="persons"></param>
        /// <param name="personsinfo"></param>
        private void updateCustomField(BSPersons persons, BSPersonsInfoFEMSA personsinfo)
        {
            bool bBIS = false;
            string id = null;
            try
            {
                this.LogTask.Information("Adicionando os campos adicionais em uma pessoa existente.");

                if (!String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "RG")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                    {
                        if ((persons = this.loadPerson(persons, personsinfo.PERSID)) != null)
                            bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "RG", "STRING", personsinfo.RG, out persons) || !bBIS;
                    }
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "RG", personsinfo.RG);
                }

                if (!String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "CPF")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "CPF", "STRING", personsinfo.CPF, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "CPF", personsinfo.CPF);
                }

                if (!String.IsNullOrEmpty(personsinfo.PISPASEP) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "PISPASEP")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "PISPASEP", "STRING", personsinfo.PISPASEP, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "PISPASEP", personsinfo.PISPASEP);
                }

                if (!String.IsNullOrEmpty(personsinfo.INTERNOEXTERNO) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "INTERNOEXTERNO")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "INTERNOEXTERNO", "STRING", personsinfo.INTERNOEXTERNO, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "INTERNOEXTERNO", personsinfo.INTERNOEXTERNO);
                }

                if (!String.IsNullOrEmpty(personsinfo.VA) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "VA")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "VA", "STRING", personsinfo.VA, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "VA", personsinfo.VA);
                }

                if (!String.IsNullOrEmpty(personsinfo.VT) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "VT")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "VT", "STRING", personsinfo.VT, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "VT", personsinfo.VT);
                }

                if (!String.IsNullOrEmpty(personsinfo.DEFICIENCIA) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "DEFICIENCIA")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)) && !String.IsNullOrEmpty(personsinfo.DEFICIENCIA))
                    {
                        if ((persons = this.loadPerson(persons, personsinfo.PERSID)) != null)
                            bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "DEFICIENCIA", "STRING", personsinfo.DEFICIENCIA, out persons) || !bBIS;
                    }
                    else if (!String.IsNullOrEmpty(personsinfo.DEFICIENCIA))
                        this.setCustomFieldSQL(personsinfo.PERSID, "DEFICIENCIA", personsinfo.DEFICIENCIA);
                }

                if (!String.IsNullOrEmpty(personsinfo.DEPARTMATTEND) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "DENGRPEMPREGADO")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "DENGRPEMPREGADO", "STRING", personsinfo.DEPARTMATTEND, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "DENGRPEMPREGADO", personsinfo.DEPARTMATTEND);
                }

                if (!String.IsNullOrEmpty(personsinfo.DATAADMISSAO) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "DATAADMISSAO")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "DATAADMISSAO", "STRING", personsinfo.DATAADMISSAO, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "DATAADMISSAO", personsinfo.DATAADMISSAO);
                }

                if (!String.IsNullOrEmpty(personsinfo.NOMERECEBEDOR) && !String.IsNullOrWhiteSpace(id = BSSQLPersons.GetIDAdditionalField(this.bisConnection, "NOMERECEBEDOR")))
                {
                    if (String.IsNullOrEmpty(BSSQLPersons.GetAdditionalFieldValue(this.bisConnection, personsinfo.PERSID, id)))
                        bBIS = this.setCustomFieldBIS(persons, personsinfo.PERSID, "NOMERECEBEDOR", "STRING", personsinfo.NOMERECEBEDOR, out persons) || !bBIS;
                    else
                        this.setCustomFieldSQL(personsinfo.PERSID, "NOMERECEBEDOR", personsinfo.NOMERECEBEDOR);
                }

                if (bBIS)
                {
                    if (persons.Update() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                        this.LogTask.Information("Campos adicionais gravados com sucesso.");
                    else
                        this.LogTask.Information(String.Format("Erro ao gravar campos adicionais. {0}", persons.GetErrorMessage()));
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
        }
        public void Import(FemsaNetworkCredential credential)
        {
            StreamReader reader = null;
            List<FemsaInfo> femsaInfo = new List<FemsaInfo>();
            int total = 0;
            FileInfo fileInfo = null;
            try
            {
                this.LogTask.Information(String.Format("Rotina iniciada as {0}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")));
                
                this.sendStartedMail(this.bisConnection, String.Format("Importacao SAP - Rotina iniciada na data {0}", DateTime.Now.ToString(new CultureInfo("pt-BR", true).DateTimeFormat.FullDateTimePattern)));

                //this.LogTask.Information("Conectando com a pasta compartilhada: {0}", credential.SharingFolder);
                //if ((reader = credential.CopyOpenFile(out fileInfo)) == null)
                //    throw new Exception("Falha ao abrir arquivo compartilhado.");
                //this.LogTask.Information("Arquivo lido com sucesso.");
                //this.LogTask.Information(String.Format("Data do arquivo: {0}", fileInfo.LastWriteTime.ToString("dd/MM/yyyy HH:mm")));
                reader = new StreamReader(@"c:\temp\Bosch_16.csv");

                string line = null;
                BSEventsInfo events = null;
                BSPersonsInfoFEMSA personsFemsa = null;
                RE_STATUS reStatus = RE_STATUS.ERROR;
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\Bosch_16.csv");

                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split('|');

                    string persno = arr[0].TrimStart(new Char[] { '0' });
                    string nome = arr[1];
                    string rg = arr[2]; //rg
                    string cpf = arr[3]; //cpf
                    string pispasep = arr[4]; //pispasep
                    string sexo = arr[5];
                    string deficiencia = arr[7];//deficiencia
                    string cnpj = arr[8];//cnpj
                    string codigoarea = arr[9];
                    string descricaoarea = arr[10];
                    string subgrupoempregado = arr[13];//departmattend
                    string unidade = arr[15];//department
                    string centrodecusto = arr[17];//costofcentre
                    string codigocentrodecusto = arr[18];//activity
                    string status = arr[19];//ativo ou 
                    string grupoempregado = arr[20];//centraloffice
                    string dataadmissao = arr[21];//dataadmissao
                    string funcao = arr[22];//job
                    string internoexterno = arr[23];//internoexterno
                    string cardno = arr[24].TrimStart(new Char[] { '0' });
                    string va = arr[25].ToLower().Equals("s") ? "1" : "0";//va
                    string vt = arr[26].ToLower().Equals("s") ? "1" : "0";//vt
                    string nomerecebedor = arr[28];//nomerecebedor

                    personsFemsa = new BSPersonsInfoFEMSA()
                    {
                        PERSNO = persno,
                        NOME = nome,
                        RG = rg,
                        CPF = cpf,
                        PISPASEP = pispasep,
                        SEX = sexo.ToLower().Equals("M") ? 0 : 1,
                        DEFICIENCIA = deficiencia,
                        CNPJ = cnpj.Replace(".", "").Replace("/", "").Replace("-", ""),
                        DEPARTMATTEND = subgrupoempregado,
                        DEPARTMENT = unidade,
                        ACTIVITY = codigocentrodecusto,
                        COSTCENTRE = centrodecusto,
                        CENTRALOFFICE = grupoempregado,
                        INTERNOEXTERNO = internoexterno,
                        STATUSPROPRIO = status,
                        VA = va,
                        VT = vt,
                        CARDNO = cardno,
                        NOMERECEBEDOR = nomerecebedor,
                        JOB = funcao,
                        DATAADMISSAO = dataadmissao,
                        CODIGOAREA = codigoarea,
                        DESCRICAOAREA = descricaoarea
                    };

                    this.LogTask.Information(String.Format("{0} de {1}", (currentline++).ToString(), totallines.ToString()));
                    this.LogTask.Information(String.Format("RE: {0}, Nome: {1}, Status: {2}", personsFemsa.PERSNO, personsFemsa.NOME, personsFemsa.STATUSPROPRIO));

                    //if (status.Equals("ativo") && !String.IsNullOrEmpty(personsFemsa.CNPJ) && personsFemsa.CNPJ.Length > 8)
                    //    this.bisConnection.executeProcedure(String.Format("update bsuser.persons set companyid = '{0}' where persno = '{1}'",(personsFemsa.CNPJ.Substring(0, 8) == "61186888" ? "0013184849D0FE94" : "00134B2266C87B31"),  persno));
                    //continue;

                    //if (status.Equals("saiu da empresa"))
                    //{
                    //    this.LogTask.Information("Pesquisando a existencia do colaborador.");
                    //    string sql = String.Format("select per.persid from bsuser.persons per left outer join bsuser.cards cd on per.persid = cd.persid where persno = '{0}' and per.status = 1 and cardid is null", personsFemsa.PERSNO);
                    //    using (DataTable table = this.bisConnection.loadDataTable(sql))
                    //    {
                    //        if (table.Rows.Count > 0)
                    //        {
                    //            foreach(DataRow row in table.Rows)
                    //            {
                    //                this.LogTask.Information("Carregado o colaborador.");
                    //                BSPersons per = new BSPersons(this.bisManager, this.bisConnection);
                    //                if (per.Load(row["persid"].ToString()))
                    //                {
                    //                    this.LogTask.Information("Colaborador carregado com sucesso.");
                    //                    if (per.Delete() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                    //                        this.LogTask.Information("Colaborador excluido com sucesso.");
                    //                    else
                    //                        this.LogTask.Information("Erro ao excluir o colaborador.");
                    //                }
                    //                else
                    //                    this.LogTask.Information("Erro ao carregar o o colaborador.");
                    //            }
                    //        }
                    //        else
                    //            this.LogTask.Information("Colaborador nao encontardo.");

                    //    }
                    //}

                    //continue;
                    IMPORT resultImport = this.existRE(personsFemsa.PERSNO, personsFemsa.NOME, personsFemsa.STATUSPROPRIO, out events);

                    if (resultImport == IMPORT.IMPORT_SAMERE)
                    {
                        personsFemsa.PERSID = events.PERSID;
                        //this.updateCustomField(new BSPersons(this.bisManager, this.bisConnection), personsFemsa);
                    }
                    else if (resultImport == IMPORT.IMPORT_NEW)
                    {
                        personsFemsa.PERSID = events.PERSID;
                        reStatus = this.save(personsFemsa);
                        femsaInfo.Add(new FemsaInfo() { Persno = personsFemsa.PERSNO, Nome = personsFemsa.NOME, Cardno = personsFemsa.CARDNO, Status = reStatus });
                    }
                    else if (resultImport == IMPORT.IMPORT_DOUBLERE)
                    {
                        personsFemsa.PERSID = events.PERSID;
                        this.updateCardidDouble(personsFemsa.PERSID, events.CARDID);
                        //this.updateCustomField(new BSPersons(this.bisManager, this.bisConnection), personsFemsa);
                    }
                    else if (resultImport == IMPORT.IMPORT_DELETE)
                    {
                        reStatus = RE_STATUS.DELETE;
                        femsaInfo.Add(new FemsaInfo() { Persno = personsFemsa.PERSNO, Nome = personsFemsa.NOME, Cardno = personsFemsa.CARDNO, Status = reStatus });
                    }
                    ++total;
                }
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

                this.LogTask.Information("Rotina finalizada.");
                this.sendFinish(this.bisConnection, femsaInfo, total);
            }
        }

        public void ImportList(List<BSPersonsInfoFEMSA> list)
        {
            StreamReader reader = null;
            List<FemsaInfo> femsaInfo = new List<FemsaInfo>();
            int total = 0;
            FileInfo fileInfo = null;
            try
            {
                this.LogTask.Information(String.Format("Rotina iniciada as {0}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")));

                reader = new StreamReader(@"c:\temp\bosch_3");

                string line = null;
                BSEventsInfo events = null;
                //BSPersonsInfoFEMSA personsFemsa = null;
                RE_STATUS reStatus = RE_STATUS.ERROR;
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                //reader = new StreamReader(@"c:\temp\boschadmitidos-fevereiro.csv");

                //while ((line = reader.ReadLine()) != null)
                foreach (BSPersonsInfoFEMSA personsFemsa in list)
                {
                    this.LogTask.Information(String.Format("{0} de {1}", (currentline++).ToString(), totallines.ToString()));
                    this.LogTask.Information(String.Format("RE: {0}", personsFemsa.PERSNO));

                    //if ((personsFemsa = list.Find(s => s.PERSNO.Equals(line))) == null)
                    //{
                    //    this.LogTask.Information("Nao encontrado.");
                    //    continue;
                    //}

                    IMPORT resultImport = this.existRE(personsFemsa.PERSNO, personsFemsa.NOME, personsFemsa.STATUSPROPRIO, out events);

                    if (resultImport == IMPORT.IMPORT_SAMERE)
                    {
                        personsFemsa.PERSID = events.PERSID;
                        //this.updateCustomField(new BSPersons(this.bisManager, this.bisConnection), personsFemsa);
                    }
                    else if (resultImport == IMPORT.IMPORT_NEW)
                    {
                        personsFemsa.PERSID = events.PERSID;
                        reStatus = this.save(personsFemsa);
                        femsaInfo.Add(new FemsaInfo() { Persno = personsFemsa.PERSNO, Nome = personsFemsa.NOME, Cardno = personsFemsa.CARDNO, Status = reStatus });
                    }
                    else if (resultImport == IMPORT.IMPORT_DOUBLERE)
                    {
                        personsFemsa.PERSID = events.PERSID;
                        this.updateCardidDouble(personsFemsa.PERSID, events.CARDID);
                        //this.updateCustomField(new BSPersons(this.bisManager, this.bisConnection), personsFemsa);
                    }
                    else if (resultImport == IMPORT.IMPORT_DELETE)
                    {
                        reStatus = RE_STATUS.DELETE;
                        femsaInfo.Add(new FemsaInfo() { Persno = personsFemsa.PERSNO, Nome = personsFemsa.NOME, Cardno = personsFemsa.CARDNO, Status = reStatus });
                    }
                    ++total;
                }
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

                this.LogTask.Information("Rotina finalizada.");
                this.sendFinish(this.bisConnection, femsaInfo, total);
            }
        }

        public void ImportSG3Excel()
        {
            List<FemsaInfo> femsaInfo = new List<FemsaInfo>();
            StreamReader reader = new StreamReader(@"c:\temp\sg3excel.csv");
            int total = 0;
            try
            {
                this.LogTask.Information(String.Format("Rotina iniciada as {0}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")));

                string line = null;
                BSEventsInfo events = null;
                BSPersonsInfoSG3 personsSG3 = null;
                RE_STATUS reStatus = RE_STATUS.ERROR;
                int currentline = 1;
                int totallines = 0;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\sg3excel.csv");
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');
                    personsSG3 = new BSPersonsInfoSG3()
                    {
                        PERSNO = arr[3],
                        NOME = arr[0],
                        RG = arr[4],
                        CPF = arr[3],
                        COMPANYNO = arr[5],
                        ESTABELECIMENTO = "CONCENTRADORA JUNDIAI"
                    };

                    this.LogTask.Information(String.Format("{0} de {1}", (currentline++).ToString(), totallines.ToString()));
                    this.LogTask.Information(String.Format("CPF: {0}, Nome: {1}", personsSG3.PERSNO, personsSG3.NOME));


                    IMPORT resultImport = this.existRE(personsSG3.PERSNO, personsSG3.NOME, personsSG3.STATUSTERCEIRO, out events);

                    if (resultImport == IMPORT.IMPORT_SAMERE)
                    {
                        personsSG3.PERSID = events.PERSID;
                        //this.updateCustomField(new BSPersons(this.bisManager, this.bisConnection), personsFemsa);
                    }
                    else if (resultImport == IMPORT.IMPORT_NEW)
                    {
                        personsSG3.PERSID = events.PERSID;
                        personsSG3.GRADE = "SG3";
                        personsSG3.PERSCLASSID = "0013B4437A2455E1";

                        reStatus = this.save(personsSG3);
                        femsaInfo.Add(new FemsaInfo() { Persno = personsSG3.PERSNO, Nome = personsSG3.NOME, Status = reStatus });
                    }
                    else if (resultImport == IMPORT.IMPORT_DOUBLERE)
                    {
                        personsSG3.PERSID = events.PERSID;
                        this.updateCardidDouble(personsSG3.PERSID, events.CARDID);
                        //this.updateCustomField(new BSPersons(this.bisManager, this.bisConnection), personsFemsa);
                    }
                    else if (resultImport == IMPORT.IMPORT_DELETE)
                    {
                        reStatus = RE_STATUS.DELETE;
                        femsaInfo.Add(new FemsaInfo() { Persno = personsSG3.PERSNO, Nome = personsSG3.NOME, Status = reStatus });
                    }

                    ++total;
                }
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

                this.LogTask.Information("Rotina finalizada.");
                this.sendFinish(this.bisConnection, femsaInfo, total);
            }
        }

        public void ImportEAB()
        {
            StreamReader reader = null;
            List<FemsaInfo> femsaInfo = new List<FemsaInfo>();
            int total = 0;
            try
            {
                this.LogTask.Information(String.Format("Rotina iniciada as {0}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")));

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.bisManager.Connect())
                    this.LogTask.Information(this.bisManager.GetErrorMessage());
                this.LogTask.Information("BIS Conectado com sucesso.");

                //this.LogTask.Information("Conectando com a pasta compartilhada: {0}", credential.SharingFolder);
                //if ((reader = credential.CopyOpenFile(out fileInfo)) == null)
                //    throw new Exception("Falha ao abrir arquivo compartilhado.");
                //this.LogTask.Information("Arquivo lido com sucesso.");
                //this.LogTask.Information(String.Format("Data do arquivo: {0}", fileInfo.LastWriteTime.ToString("dd/MM/yyyy HH:mm")));
                reader = new StreamReader(@"c:\temp\eab");

                string line = null;
                BSEventsInfo events = null;
                BSPersonsInfoFEMSA personsFemsa = null;
                RE_STATUS reStatus = RE_STATUS.ERROR;
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\eab");

                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string nome = arr[0];
                    string tipopessoa = arr[1];
                    string email = arr[2];
                    string telefone = arr[3];
                    string authid = arr[4];
                    string cardno = arr[5];
                    string persno = arr[6].TrimStart(new Char[] { '0' });

                    personsFemsa = new BSPersonsInfoFEMSA()
                    {
                        PERSNO = persno,
                        NOME = nome,
                        EMAIL = email,
                        PHONEPRIVATE = telefone,
                        AUTHID = authid,
                        CARDNO = cardno,
                        PERSCLASSID = tipopessoa
                    };

                    this.LogTask.Information(String.Format("{0} de {1}", (currentline++).ToString(), totallines.ToString()));
                    this.LogTask.Information(String.Format("RE: {0}, Nome: {1}, N.Cartao: {2}", personsFemsa.PERSNO, personsFemsa.NOME, !String.IsNullOrEmpty(personsFemsa.CARDNO) ? personsFemsa.CARDNO : "Sem Cartao"));

                    BSPersons persons = new BSPersons(this.bisManager, this.bisConnection);
                    string firstname = null;
                    string lastname = null;

                    BSCommon.ParseName(personsFemsa.NOME, out firstname, out lastname);
                    personsFemsa.FIRSTNAME = firstname;
                    personsFemsa.LASTNAME = lastname;

                    this.LogTask.Information("Carregando o objeto BSPersons.");
                    persons = (BSPersons)this.FillBISObject(persons, personsFemsa);
                        this.LogTask.Information("Nova pessoa.");

                    if ((String.IsNullOrEmpty(personsFemsa.PERSID) || (!String.IsNullOrEmpty(persons.GetPersonId())) ? persons.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS : false))
                    {
                        this.LogTask.Information(String.Format("Registro gravado com sucesso. PERSID: {0}", persons.GetPersonId()));
                        personsFemsa.PERSID = persons.GetPersonId();
                    }
                    else
                        throw new Exception(String.Format("Erro ao gravar o registro. {0}", persons.GetErrorMessage()));

                    if (!String.IsNullOrEmpty(personsFemsa.CARDNO))
                    {
                        this.LogTask.Information(String.Format("Atribuindo o cartao {0}", personsFemsa.CARDNO));
                        this.LogTask.Information("Verificando a existencia do cartao.");
                        string cardid = BSSQLCards.GetID(this.bisConnection, personsFemsa.CARDNO, "0", BS_STATUS.ACTIVE);
                        if (String.IsNullOrEmpty(cardid))
                        {
                            this.LogTask.Information("Criando o novo cartao.");
                            BSCards cards = new BSCards(this.bisManager);
                            cards.CARDNO = personsFemsa.CARDNO;
                            if (cards.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                            {
                                this.LogTask.Information("Cartao criado com sucesso.");
                                this.LogTask.Information("Verificando o ID do cartao.");
                                cardid = BSSQLCards.GetID(this.bisConnection, personsFemsa.CARDNO, "0", BS_STATUS.ACTIVE);
                                this.LogTask.Information("ID encontrado: {0}. Atribuindo o cartao.", cardid);
                                if (persons.AttachCard(cardid) == BS_ADD.SUCCESS)
                                    this.LogTask.Information("Cartao atribuido com sucesso.");
                                else
                                    throw new Exception("Erro ao atribuir o cartao.");
                            }
                            else
                                throw new Exception(String.Format("Erro ao criar o cartao. {0}", cards.GetErrorMessage()));
                        }
                    }
                    ++total;
                }
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

                this.LogTask.Information("Rotina finalizada.");
                this.sendFinish(this.bisConnection, femsaInfo, total);
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
        /// <param name="ptbr">Localização.</param>
        public FEMSAImport(ILogger logtask, ILogger logresult, SQLParameters bissql, SQLParameters bislogsql, BISManager bismanager)
        {
            this.LogTask = logtask;
            this.LogResult = logresult;
            this.bisConnection = new HzConexao(bissql.SQLHost, bissql.SQLUser, bissql.SQLPwd, "acedb", "System.Data.SqlClient");
            this.bisLogConnection = new HzConexao(bislogsql.SQLHost, bislogsql.SQLUser, bislogsql.SQLPwd, "BISEventLog", "System.Data.SqlClient");
            this.bisManager = bismanager;
        }
        #endregion    
    }
}
