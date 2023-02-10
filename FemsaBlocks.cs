using AMSLib;
using AMSLib.ACE;
using AMSLib.Manager;
using AMSLib.SQL;
using HzLibConnection.Data;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FemsaTools
{
    #region Enumeration
    public enum BLACTION
    {
        OLDBLOCK = 0,
        NEWBLOCK = 1,
        UNBLOCKED = 2,
        FREEACCESS = 3,
        ERRO = 4,
        NOTFOUND = 5,
        WRONGTIME = 6,
        NOCARD = 7,
        JURBATUBACORPORATIVO = 8,
        NONE = 9,
        ERROBLOCK = 10,
        ERROUNBLOCK = 11
    }
    #endregion

    public class BSResult
    {
        #region variables
        /// <summary>
        /// Dia da verificação.
        /// </summary>
        public string Data { get; set; }
        /// <summary>
        /// Matrícula da Pessoa.
        /// </summary>
        public string RE { get; set; }
        /// <summary>
        /// Nome da pessoa.
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Tipo da ação.
        /// </summary>
        public BLACTION Action { get; set; }
        /// <summary>
        /// Mensagem de status.
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// Data da entrada no BIS.
        /// </summary>
        public string DataEntradaBIS { get; set; }
        /// <summary>
        /// Data da saída no BIS.
        /// </summary>
        public string DataSaidaBIS { get; set; }
        /// <summary>
        /// Data da entrada no WorkForce.
        /// </summary>
        public string DataEntradaWFM { get; set; }
        /// <summary>
        /// Código da Situação no WorkForce.
        /// </summary>
        public string CodSituacao { get; set; }
        /// <summary>
        /// Status no WorkForce,
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// Tolerância para o horário de entrada.
        /// </summary>
        public int Tolerancia { get; set; }
        /// <summary>
        /// Site do colaborador.
        /// </summary>
        public string Site { get; set; }
        #endregion

        #region Functions
        /// <summary>
        /// Grava as informações dos resultados das ações de integracao no banco de dados.
        /// </summary>
        /// <param name="connection">Conexão com o SQL Server.</param>
        public void Save(HzConexao connection)
        {
            try
            {
                string sql = String.Format("set dateformat 'dmy' exec Horizon..spIncluirWFMBIS '{0}', '{1}', '{2}', '{3}', '{4}', {5}, {6}, {7}, '{8}', '{9}', {10}, '{11}'",
                    this.Data, this.RE, this.Name.Replace("'", ""), this.Action.ToString(), (this.Message == null ? "null" : this.Message.Replace("'", "")), (this.DataEntradaBIS == null ? "null" : "'" + this.DataEntradaBIS + "'"),
                    (this.DataSaidaBIS == null ? "null" : "'" + this.DataSaidaBIS + "'"), (this.DataEntradaWFM == null ? "null" : "'" + this.DataEntradaWFM + "'"),
                    this.CodSituacao, this.Status, this.Tolerancia.ToString(), this.Site);
                connection.executeProcedure(sql);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion
    }

    /// <summary>
    /// Classe com as funcionalidade de bloqueio.
    /// </summary>
    public class FemsaBlocks
    {
        #region Variables
        /// <summary>
        /// Conexão com o BIS.
        /// </summary>
        public BISManager bisManager { get; set; }
        /// <summary>
        /// Conexão com o banco de dados de Eventos do BIS.
        /// </summary>
        public HzConexao bisConnection { get; set; }
        /// <summary>
        /// Conexão com o banco de dados de BIS.
        /// </summary>
        public HzConexao bisACEConnection { get; set; }
        /// <summary>
        /// Log da aplicação.
        /// </summary>
        public ILogger LogTask { get; set; }
        /// <summary>
        /// Log dos resultados.
        /// </summary>
        public ILogger LogResult { get; set; }
        /// <summary>
        /// Localização do serivdor.
        /// </summary>
        public CultureInfo ptBR { get; set; }

        #endregion   

        #region EMAIL
        /// <summary>
        /// Envia e-mail.
        /// </summary>
        /// <param name="subject">Título do e-mail.</param>
        /// <param name="body">Corpo do e-mail.</param>
        public void SendMail(string subject, string body)
        {
            try
            {
                string recipients = "'ronaldo.moscoso@grupoorion.com.br;ronaldomoscoso@gmail.com'";
                string sql = String.Format("exec msdb..sp_send_dbmail @profile_name = 'Femsa', @recipients = {0}, @subject = '{1}', @body = '{2}', @body_format = 'html'",
                    recipients, subject, body);
                //karen.costa@kof.com.mx;fernanda.santos@kof.com.mx
                //this.bisConnection.executeProcedure(sql);
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
        public void SendStartedMail()
        {
            try
            {
                this.LogTask.Information("Enviando e-mail do inicio da integracao.");

                string subject = String.Format("integracao WFM x BIS - Rotina iniciada na data {0}", DateTime.Now.ToString(ptBR.DateTimeFormat.FullDateTimePattern));
                string HTML = "<html><body>Rotina Iniciada com Sucesso!</body></html>";
                this.SendMail(subject, HTML);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }

        /// <summary>
        /// Envia e-mail com os resultados da integracao.
        /// </summary>
        /// <param name="data">Data da rotina.</param>
        /// <param name="results">Coleção com os resultados da rotina de integracao.</param>
        public void SendFinishedMail(string data, List<BSResult> results)
        {
            try
            {
                this.LogTask.Information("Preparando e-mail com as informacoes da integracao WFM x BIS");

                string subject = String.Format("integracao WFM x BIS - Rotina executada na data {0}", DateTime.Now.ToString(ptBR.DateTimeFormat.FullDateTimePattern));
                string body = results.Count > 0 ? String.Format("Processamento realizado com sucesso. {0} pessoas processadas", results.Count.ToString()) :
                    "Erro ao realizar a integracao WFM x BIS!";
                List<BSResult> resultNF = results.FindAll(d => d.Action == BLACTION.NOTFOUND || d.Action == BLACTION.NOCARD);
                string HTML = "<html><body>Rotina Executada com Sucesso!</body></html>";
                this.SendMail(subject, HTML);
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
        public void SendErrorMail(List<BSResult> results)
        {
            try
            {
                this.LogTask.Information("Erro na rotina de execucao");

                List<BSResult> errors = results.FindAll((delegate (BSResult b) { return b.Action == BLACTION.ERRO || b.Action == BLACTION.ERROBLOCK || b.Action == BLACTION.ERROUNBLOCK; }));
                string subject = "Integracao WFM x BIS - Erro na rotina de execucao";
                string body = "<table>";
                string msg = null;
                try
                {
                    foreach (BSResult r in errors)
                    {
                        if (r.Action == BLACTION.ERROBLOCK)
                            msg = "Erro ao Bloquear";
                        else if (r.Action == BLACTION.ERROUNBLOCK)
                            msg = "Erro ao desbloquear";
                        else
                            msg = "Erro Generico";


                        body += String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>", msg, r.Site, r.RE, r.Name);
                    }
                    body += "</table>";
                }
                catch
                {

                }
                string HTML = String.Format("<html><body>{0}</body></html>", body);
                this.SendMail(subject, HTML);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
            }
        }
        #endregion

        #region BLOCK
        /// <summary>
        /// Verifica as exceções para a rotina de bloqueio.
        /// </summary>
        /// <returns></returns>
        public bool CheckExceptions(BISManager bisManager)
        {
            try
            {
                int iCount = -1;
                BSPersons persons = null;
                bool bErro = false;
                List<BSResult> results = new List<BSResult>();
                DateTime startdate = DateTime.Now;

                if (startdate.Hour >= 20)
                    startdate = startdate.AddDays(1);

                this.SendMail("Rotina de Execao", "Inicio da Rotina de Excecoes.");

                this.LogTask.Information("Verificando as excecoes.");
                string sql = String.Format("set dateformat 'dmy' select per.persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), cmpDtInicio, cmpDtTermino, " +
                    "lockid, cmpVlTolerancia from hzRH.dbo.tblBlockExcecao exce inner join bsuser.persons per on per.persid = exce.persid inner join Horizon..tblTolerancia tol on tol.clientid = per.clientid " +
                    "left outer join bsuser.lockouts lock on lock.persid = per.persid where cmpDtInicio >= '{0} 00:00:00' and cmpDtInicio <= '{0} 23:59:59'", startdate.ToString("dd/MM/yyyy"));
                using (DataTable tablExce = bisACEConnection.loadDataTable(sql))
                {
                    if (tablExce != null || tablExce.Rows.Count > 0)
                    {
                        this.LogTask.Information(String.Format("Registros encontrados: {0}", tablExce.Rows.Count.ToString()));
                        this.SendMail("Rotina de Excecao", String.Format("Registros encontrados: {0}", tablExce.Rows.Count.ToString()));
                        iCount = 1;
                        foreach (DataRow rowExce in tablExce.Rows)
                        {
                            this.LogTask.Information("Verificando no BIS {0} de {1}. RE: {2}, Nome: {3}, status: {4}, horario_inicio: {5}, horario_termino: {6}", (iCount++).ToString(), tablExce.Rows.Count.ToString(),
                                rowExce["persno"].ToString(), rowExce["nome"].ToString(), String.IsNullOrEmpty(rowExce["lockid"].ToString()) ? "Nao ha bloqueio." : "bloqueado.", rowExce["cmpDtInicio"].ToString(), rowExce["cmpDtTermino"].ToString());

                            int tolerancia = 40;
                            if (!int.TryParse(rowExce["cmpVlTolerancia"].ToString(), out tolerancia))
                                tolerancia = 40;

                            BSResult result = new BSResult();
                            result.Data = rowExce["cmpDtInicio"].ToString();
                            result.RE = rowExce["persno"].ToString();
                            result.Name = rowExce["nome"].ToString();
                            result.Action = BLACTION.NONE;
                            result.Tolerancia = tolerancia;
                            result.Message = "Rotina de Excecao";

                            if (persons == null)
                                persons = new BSPersons(bisManager, bisACEConnection);
                            this.LogTask.Information("Carregando a pessoa");
                            if (!persons.Load(rowExce["persid"].ToString()))
                            {
                                this.LogTask.Information("Erro ao carregar a pessoa.");
                                this.SendMail("Rotina de Excecoes", "Erro de processamento no BIS. Rotina de Excecao interrompida! Erro Severo! Favor contactar o suporte!");
                                break;
                            }
                            this.LogTask.Information("Pessoa carregada com sucesso.");

                            if (!String.IsNullOrEmpty(rowExce["lockid"].ToString()))
                            {
                                this.LogTask.Information("Desbloqueando a pessoa.");
                                if (persons.RemoveAllBLocks(bisACEConnection, rowExce["persid"].ToString()) == HzBISCommands.BS_ADD.ERROR)
                                {
                                    result.Action = BLACTION.ERROUNBLOCK;
                                    bErro = true;
                                    this.LogTask.Information("Erro ao desbloquear a pessoa. Verificar o desbloqueio.");
                                    continue;
                                }
                                this.LogTask.Information("Pessoa desbloqueada com sucesso.");
                            }

                            DateTime dtExce = DateTime.Parse(rowExce["cmpDtInicio"].ToString());
                            DateTime dtExceEnd = DateTime.Parse(rowExce["cmpDtTermino"].ToString());
                            DateTime horario_ent = new DateTime(dtExce.Year, dtExce.Month, dtExce.Day, dtExce.Hour, dtExce.Minute, dtExce.Second);
                            DateTime horario_sai = new DateTime(dtExceEnd.Year, dtExceEnd.Month, dtExceEnd.Day, dtExceEnd.Hour, dtExceEnd.Minute, dtExceEnd.Second);
                            persons.AUTHFROM = (new DateTime(horario_ent.Year, horario_ent.Month, horario_ent.Day, horario_ent.Hour, horario_ent.Minute, 0)).AddMinutes(tolerancia * (-1));
                            persons.AUTHUNTIL = new DateTime(horario_sai.Year, horario_sai.Month, horario_sai.Day, horario_sai.Hour, horario_sai.Minute, 0);
                            this.LogTask.Information(String.Format("Atualizando a validade da pessoa para a data de entrada: {0} e saida: {1}", horario_ent.ToString(ptBR.DateTimeFormat.ShortDatePattern + " HH:mm"),
                                horario_sai.ToString(this.ptBR.DateTimeFormat.ShortDatePattern + " HH:mm")));
                            if (persons.Save() == HzBISCommands.BS_SAVE.SUCCESS)
                                this.LogTask.Information("Validade atualizada com suceso.");
                            else
                            {
                                result.Action = BLACTION.ERRO;
                                this.LogTask.Information(String.Format("Erro ao atualizar a validade. Erro: {0}", persons.GetErrorMessage()));
                                bErro = true;
                            }
                        }

                        if (bErro)
                            this.SendErrorMail(results);
                        this.SendMail("Rotina de Excecoes", "Rotina de excecoes terminada com sucesso.");
                    }
                    else
                    {
                        this.LogTask.Information("Nao ha excecoes cadastradas.");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Executa os bloqueios/desbloqueios das pessoas.
        /// </summary>
        /// <param name="trigger">Objeto para o intervalo de gravação no BIS.</param>
        /// <returns></returns>
        public bool CheckBlocks(string re)
        {
            try
            {
                DateTime startdate = DateTime.Now.Hour < 20 ? DateTime.Now : DateTime.Now.AddDays(1);
                int tolerancia = 40;
                List<BSResult> results = new List<BSResult>();
                bool bErro = false;
                int bErroLoad = 0;
                int bErroSave = 0;
                int countRE = 1;

                this.LogTask.Information("Iniciando a rotina de validacao do acesso atraves do WFM!");
                if (String.IsNullOrEmpty(re))
                    this.SendStartedMail();
                int iCount = 1;
                BSPersons persons = null;
                string sql = null;

                this.LogTask.Information("Conectando com o BIS.");
                if (this.bisManager.Connect())
                {
                    this.LogTask.Information("BIS conectado com sucesso.");
                    DataTable tableWF = null;
                    if ((tableWF = this.RunQuery(startdate, re)) != null)
                    {
                        if (tableWF != null && tableWF.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("Registros encontrados: {0}", tableWF.Rows.Count.ToString()));
                            this.SendMail("Rotina de Bloqueio/Desbloqueio", String.Format("Registros encontrados: {0}", tableWF.Rows.Count.ToString()));
                            bool bErroBIS = false;

                            foreach (DataRow rowWF in tableWF.Rows)
                            {
                                this.LogTask.Information("Verificando no BIS {0} de {1}. RE: {2}, Nome: {3}, status: {4}, cod_situacao: {5}, horario_ent: {6}", (iCount++).ToString(), tableWF.Rows.Count.ToString(),
                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), rowWF["status"].ToString(), rowWF["cod_situacao"].ToString(), !String.IsNullOrEmpty(rowWF["horario_ent"].ToString()) ? rowWF["horario_ent"].ToString() : "Nao ha horario de entrada.");

                                sql = String.Format("select distinct per.persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), lockid, lock.CAUSEOFLOCK, authfrom = convert(varchar, authfrom, 103) + ' ' + convert(varchar, authfrom, 108), " +
                                    "cmpVlTolerancia, per.clientid from bsuser.persons per inner join Horizon..tblTolerancia tol on tol.clientid = per.clientid left outer join bsuser.cards cd on cd.persid = per.persid " +
                                    "left outer join bsuser.LOCKOUTS lock on per.persid = lock.persid inner join bsuser.acpersons acpe on acpe.persid = per.persid " +
                                    "where persno = '{0}' and per.status = 1 and tol.cmpInStatus = 1 and cd.status != 0 and persclass = 'E' ", rowWF["re"].ToString());

                                if (iCount == tableWF.Rows.Count / 2)
                                    this.SendMail("Rotina de Bloqueio/Desbloqueio", "50% dos registros processados!");

                                using (DataTable tableBIS = bisACEConnection.loadDataTable(sql))
                                {
                                    try
                                    {
                                        // Executa os REs duplicados.
                                        
                                        BSResult result = new BSResult();
                                        result.Data = rowWF["data"].ToString();
                                        result.RE = rowWF["re"].ToString();
                                        result.Name = rowWF["nome"].ToString();
                                        result.Action = BLACTION.NONE;
                                        result.Site = rowWF["site"].ToString();
                                        result.Tolerancia = tolerancia;
                                        result.Message = rowWF["descricao_situacao"].ToString();

                                        if (tableBIS == null || tableBIS.Rows.Count < 1)
                                        {
                                            result.Action = BLACTION.NOTFOUND;
                                            results.Add(result);
                                            this.LogTask.Information("Pessoa nao encontrada.");
                                            continue;
                                        }

                                        countRE = 1;
                                        foreach (DataRow rowBIS in tableBIS.Rows)
                                        {
                                            if (tableBIS.Rows.Count > 1)
                                                this.LogTask.Information(String.Format("{0} de {1} RE duplicados.", (countRE++).ToString(), tableBIS.Rows.Count.ToString()));

                                            //if (String.IsNullOrEmpty(rowBIS["cardid"].ToString()))
                                            //{
                                            //    result.Action = BLACTION.NOCARD;
                                            //    results.Add(result);
                                            //    this.LogTask.Information("Pessoa sem cartao.");
                                            //    continue;
                                            //}

                                            if (!int.TryParse(rowBIS["cmpVlTolerancia"].ToString(), out tolerancia))
                                                tolerancia = 40;

                                            bool isBlocked = !String.IsNullOrEmpty(rowBIS["lockid"].ToString());
                                            bool isAllowedRestrictAccess = (rowWF["cod_situacao"].ToString().Equals("3000") || rowWF["cod_situacao"].ToString().Equals("220") ||
                                                rowWF["cod_situacao"].ToString().Equals("225") || rowWF["cod_situacao"].ToString().Equals("800"));
                                            bool isAllowedFreeAccess = rowWF["cod_situacao"].ToString().Equals("110");

                                            this.LogTask.Information("Pessoa encontrada. Persid: {0}, lockid: {1}", rowBIS["persid"].ToString(),
                                                String.IsNullOrEmpty(rowBIS["lockid"].ToString()) ? "Nao ha bloqueio" : rowBIS["lockid"].ToString() + " / Motivo: " + rowBIS["causeoflock"].ToString());

                                            result.DataEntradaWFM = !String.IsNullOrEmpty(rowWF["horario_ent"].ToString()) ? String.Format("{0} {1}", DateTime.Parse(rowWF["horario_ent"].ToString()).ToShortDateString(),
                                            DateTime.Parse(rowWF["horario_ent"].ToString()).ToShortTimeString()) : "Nao ha horario de entrada.";
                                            result.CodSituacao = rowWF["cod_situacao"].ToString();
                                            result.Status = rowWF["status"].ToString();

                                            //se a pessoa já está bloqueada no BIS e seu status continua bloquado, ignora.
                                            if (isBlocked && rowWF["status"].ToString().Equals("1"))
                                            {
                                                result.Action = BLACTION.OLDBLOCK;
                                                results.Add(result);

                                                this.LogTask.Information("Bloqueio Existente. Nenhuma acao realizada!");
                                                this.LogResult.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "PESSOA JA BLOQUEADA. NENHUMA ACAO."));
                                                continue;
                                            }

                                            if (persons == null)
                                                persons = new BSPersons(this.bisManager, this.bisACEConnection);
                                            
                                            this.LogTask.Information("Carregando a pessoa");
                                            if (!persons.Load(rowBIS["persid"].ToString()))
                                            {
                                                result.Action = BLACTION.ERRO;
                                                results.Add(result);

                                                bErroBIS = true;
                                                this.LogTask.Information(String.Format("Erro de processamento no BIS. PERSID: {0}. Erro: {1}", rowBIS["persid"].ToString(),
                                                    persons.GetErrorMessage()));
                                                ++bErroLoad;

                                                if (bErroLoad == 5 && String.IsNullOrEmpty(re))
                                                    this.SendMail("Rotina de Bloqueio/Desbloqueio", "Houve mais de 5 erros ao carregar o colaborador. Favor entrar em contato com o suporte!");
                                            }
                                            this.LogTask.Information("Pessoa carregada com sucesso.");

                                            // se a pessoa já está bloqueada e seu status é desbloqueio, desbloqueia a pessoa.
                                            if (isBlocked && rowWF["status"].ToString().Equals("0"))
                                            {
                                                this.LogResult.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "DESBLOQUEADO"));

                                                this.LogTask.Information("Desbloqueando a pessoa.");
                                                using (DataTable tableLock = BSSQLPersons.GetAllLockIDs(bisACEConnection, rowBIS["persid"].ToString()))
                                                {
                                                    if (tableLock != null && tableLock.Rows.Count > 0)
                                                    {
                                                        this.LogTask.Information(String.Format("Ha {0} bloqueios.", tableLock.Rows.Count.ToString()));
                                                        int iLock = 1;
                                                        foreach (DataRow rowLock in tableLock.Rows)
                                                        {
                                                            this.LogTask.Information(String.Format("Removendo bloqueio {0} de {1}", (iLock++).ToString(), tableLock.Rows.Count.ToString()));
                                                            if (persons.RemoveBlock(rowLock["lockid"].ToString()) == HzBISCommands.BS_ADD.SUCCESS)
                                                            {
                                                                this.LogTask.Information("Pessoa desbloqueada com sucesso.");
                                                                result.Action = BLACTION.UNBLOCKED;
                                                                results.Add(result);
                                                            }
                                                            else
                                                            {
                                                                result.Action = BLACTION.ERROUNBLOCK;
                                                                results.Add(result);

                                                                this.LogTask.Information("Erro ao desbloquear a pessoa. Verificar o desbloqueio.");
                                                                bErro = true;
                                                            }
                                                        }
                                                    }
                                                    else
                                                        this.LogTask.Information("Nao ha bloqueios para o colaborador.");
                                                }
                                            }
                                            // se a pessoa já está desbloqueada e seu status é bloqueio e não há exceção, bloqueia a pessoa.
                                            else if (!isBlocked && rowWF["status"].ToString().Equals("1"))
                                            {
                                                this.LogResult.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "BLOQUEADO"));

                                                //bloqueia a pessoa.
                                                if (persons.AddBlock(startdate, startdate.AddDays(1), rowWF["descricao_situacao"].ToString()) == HzBISCommands.BS_ADD.SUCCESS)
                                                {
                                                    this.LogTask.Information("Pessoa bloqueada com sucesso.");
                                                    result.Message = rowWF["descricao_situacao"].ToString();
                                                    result.Action = BLACTION.NEWBLOCK;
                                                }
                                                else
                                                {
                                                    this.LogTask.Information("Erro ao bloquear a pessoa. Verificar o bloqueio.");
                                                    bErro = true;
                                                    result.Action = BLACTION.ERROBLOCK;
                                                }

                                                results.Add(result);
                                                continue;
                                            }
                                            // se a pessoa já está bloqueada e há exceção, desbloqueia a pessoa.
                                            else if (isBlocked)
                                            {
                                                this.LogResult.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "DESBLOQUEADO"));

                                                this.LogTask.Information("Desbloqueando a pessoa.");
                                                if (persons.RemoveAllBLocks(bisACEConnection, rowBIS["persid"].ToString()) == HzBISCommands.BS_ADD.ERROR)
                                                {
                                                    result.Action = BLACTION.ERROUNBLOCK;
                                                    results.Add(result);

                                                    this.LogTask.Information("Erro ao desbloquear a pessoa. Verificar o desbloqueio.");
                                                    bErro = true;
                                                    continue;
                                                }

                                                this.LogTask.Information("Pessoa desbloqueada com sucesso.");
                                                result.Action = BLACTION.UNBLOCKED;
                                            }

                                            DateTime horario_ent = !String.IsNullOrEmpty(rowWF["horario_ent"].ToString()) ? DateTime.Parse(rowWF["horario_ent"].ToString()).AddMinutes(tolerancia * (-1)) :
                                                DateTime.Now;
                                            DateTime horario_sai = !String.IsNullOrEmpty(rowWF["horario_ent"].ToString()) ? DateTime.Parse(rowWF["horario_ent"].ToString()).AddHours(18) :
                                                DateTime.Now.AddHours(18);

                                            //se for em MS, aumenta 1 hora devido ao fuso horário
                                            if (rowBIS["clientid"].ToString().Equals("001321A28697C946"))
                                                horario_ent = horario_ent.AddHours(1);

                                            if (isAllowedFreeAccess && persons.AUTHUNTIL.Year < 2000)
                                            {
                                                result.Action = BLACTION.FREEACCESS;
                                                result.DataEntradaBIS = null;
                                                result.DataSaidaBIS = null;
                                                results.Add(result);

                                                this.LogTask.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "Pessoa com acesso liberado ja cadastrado! Nenhuma acao realiada."));
                                                this.LogResult.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "Pessoa com acesso liberado ja cadastrado! Nenhuma acao realiada."));
                                            }
                                            else
                                            {
                                                this.LogResult.Information(String.Format("Persid: {0}, RE: {1}, Nome: {2}, Resultado: {3}", rowBIS["persid"].ToString(),
                                                    rowWF["re"].ToString(), rowWF["nome"].ToString(), "HORARIO ALTERADO: " + horario_ent.ToString(ptBR.DateTimeFormat.ShortDatePattern + " HH:mm")));

                                                try
                                                {
                                                    persons.AUTHFROM = new DateTime(horario_ent.Year, horario_ent.Month, horario_ent.Day, horario_ent.Hour, horario_ent.Minute, 0);
                                                    result.DataEntradaBIS = String.Format("{0} {1}", horario_ent.ToShortDateString(), horario_ent.ToShortTimeString());
                                                    result.DataSaidaBIS = null;
                                                    if (isAllowedRestrictAccess && !isAllowedFreeAccess)
                                                    {
                                                        persons.AUTHUNTIL = new DateTime(horario_sai.Year, horario_sai.Month, horario_sai.Day, horario_sai.Hour, horario_sai.Minute, 0);
                                                        result.DataSaidaBIS = String.Format("{0} {1}", horario_sai.ToShortDateString(), horario_sai.ToShortTimeString());
                                                    }
                                                    else
                                                        persons.AUTHUNTIL = DateTime.MinValue;

                                                    this.LogTask.Information(String.Format("Atualizando a validade da pessoa para a data de entrada: {0} e saida: {1}", horario_ent.ToString(ptBR.DateTimeFormat.ShortDatePattern + " HH:mm"),
                                                        horario_sai.ToString(ptBR.DateTimeFormat.ShortDatePattern + " HH:mm")));

                                                    results.Add(result);//retirar

                                                    if (persons.Save() == HzBISCommands.BS_SAVE.SUCCESS)
                                                    {
                                                        results.Add(result);
                                                        this.LogTask.Information("Validade atualizada com suceso.");
                                                    }
                                                    else
                                                    {
                                                        result.Action = BLACTION.ERRO;
                                                        results.Add(result);
                                                        this.LogTask.Information("Erro ao atualizar a validade.");
                                                        bErro = true;
                                                        bErroSave++;
                                                        if (bErroSave == 5)
                                                            this.SendMail("Rotina de Bloqueio/Desbloqueio", "Houve mais de 5 erros ao salvar a validade do colaborador. Favor entrar em contato com o suporte!");
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    this.LogTask.Information(String.Format("Erro ao processar: {0}", ex.Message));
                                                    this.SendMail("Rotina de Bloqueio/Desbloqueio", "Erro na gravação da validade! Erro Severo! Favor contactar o suporte!");
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        this.LogTask.Information(String.Format("Erro ao processar: {0}", ex.Message));
                                        if (String.IsNullOrEmpty(re))
                                            this.SendMail("Rotina de Bloqueio/Desbloqueio", "Erro de processamento no BIS. Rotina interrompida! Erro Severo! Favor contactar o suporte!");
                                    }
                                }
                            }

                            if (String.IsNullOrEmpty(re))
                                this.SendFinishedMail("", results);
                            if (bErro && String.IsNullOrEmpty(re))
                                this.SendErrorMail(results);

                            if (String.IsNullOrEmpty(re))
                                this.CheckExceptions(this.bisManager);
                        }
                        else
                        {
                            this.LogTask.Information(String.Format("Erro: {0}", this.bisManager.GetErrorMessage()));
                            if (String.IsNullOrEmpty(re))
                                this.SendMail("Rotina de Bloqueio/Desbloqueio", "Erro de processamento do WFM. Rotina interrompida! Erro Severo! Favor contactar o suporte!");
                        }
                    }
                    else
                    {
                        this.LogTask.Information("Erro ao executar a pesquisa no WFM. Nenhum reigstro retornado!");
                        if (String.IsNullOrEmpty(re))
                            this.SendMail("Rotina de Bloqueio/Desbloqueio", "Erro de processamento do WFM. Rotina interrompida! Erro Severo! Favor contactar o suporte!");
                    }
                }
                else
                {
                    this.LogTask.Information(String.Format("Integracao nao realziada devido a erro de conexao com o BIS! Erro: {0}",
                        this.bisManager.GetErrorMessage()));
                    if (String.IsNullOrEmpty(re))
                        this.SendMail("Rotina de Bloqueio/Desbloqueio", "Erro de processamento do WFM. Rotina interrompida! Erro Severo! Favor contactar o suporte!");
                }

                return true;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
                if (String.IsNullOrEmpty(re))
                    this.SendMail("Rotina de Bloqueio/Desbloqueio", "Erro Generico. Rotina interrompida! Erro Severo! Favor contactar o suporte!");
                return false;
            }
            finally
            {
                this.LogTask.Information("Desconectando com o BIS.");
                if (this.bisManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
            }
        }
        #endregion

        #region WFM
        /// <summary>
        /// Executa a query com o WFM.
        /// Caso haja algum erro na cláusula "where" que divide a carga no BIS,
        /// a instrução que verifica todo o dia é executada.
        /// Tratamento de exceção desta forma para evitar que a rotina não seja executada.
        /// </summary>
        /// <param name="connection">Conexão com o banco de dados.</param>
        /// <param name="startdate">Data da verificação do WFM.</param>
        /// <returns></returns>
        public DataTable RunQuery(DateTime startdate, string re)
        {
            try
            {
                string sqlMain = String.Format("select distinct data, be.re, be.nome, d30.cod_situacao, status, descricao_situacao, site, horario_ent, horario_sai from [10.153.68.133].sisqualwfm.[dbo].[BOSCH_EmpregadosSituacao_D+30] d30 " +
                    "inner join [10.153.68.133].sisqualwfm.dbo.BOSCH_Empregados be on be.re = d30.re inner join [10.153.68.133].sisqualwfm.dbo.BOSCH_TIPOS_SITUACAO bs on d30.COD_SITUACAO = bs.cod_situacao " +
                    "where convert(date, data, 103) = '{0}' and site not in ('JURUBATUBA', 'JURUBATUBA DIRETORIA', 'JURUBATUBA OPERAÇÕES') ", startdate.ToString("MM/dd/yyyy"));

                string sql = sqlMain;
                if (String.IsNullOrEmpty(re))
                {
                    if (startdate.Hour >= 20)
                        sql += " and((convert(time, horario_ent) >= '00:00:00' and convert(time, horario_ent) <= '12:00:00') or horario_ent is null or bs.COD_SITUACAO = 110) order by horario_ent";
                    else
                        // apenas verifica os ativos pois os bloqueados já foram analisados na execução anterior.
                        sql += " and (convert(time, horario_ent) >= '12:00:01' or horario_ent is null) and d30.cod_situacao in (3000, 110) and status in (0) order by horario_ent";
                }
                else
                {
                    sql += String.Format(" and be.RE = '{0}'", re);
                }
                this.LogTask.Information("Pesquisando no SQL SERVER...");
                this.LogTask.Information("SQL: {0}", sql);

                DataTable table = null;
                for (int i = 0; i < 3; ++i)
                {
                    try
                    {
                        table = this.bisACEConnection.loadDataTable(sql);
                        break;
                    }
                    catch (Exception ex)
                    {
                        this.LogTask.Information(ex, ex.Message);
                        Thread.Sleep(5 * 10 * 1000);
                    }
                }

                if (table == null || table.Rows.Count < 1)
                    throw new Exception("Falha na conexao com o WFM apos 3 tentativas. Favor procurar o suporte!");
                return table;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(ex, ex.Message);
                this.SendMail("Erro Conexão WFM", ex.Message);
                return null;
            }
        }
        #endregion

        #region EVENTS
        /// <summary>
        /// Construtor da classe.
        /// </summary>
        /// <param name="logtask">Log dos eventos.</param>
        /// <param name="logresult">Log dos resultados.</param>
        /// <param name="bisacesql">Parâmetros do SQL Server de configuração do BIS.</param>
        /// <param name="bissql">Parâmetros do SQL Server de evento do BIS.</param>
        /// <param name="bismanager">Objeto do BIS.</param>
        /// <param name="ptbr">Localização.</param>
        public FemsaBlocks(ILogger logtask, ILogger logresult, SQLParameters bisacesql, SQLParameters bissql, BISManager bismanager, CultureInfo ptbr)
        {
            this.ptBR = ptbr;
            this.LogTask = logtask;
            this.LogResult = logresult;
            this.bisACEConnection = new HzConexao(bisacesql.SQLHost, bisacesql.SQLUser, bisacesql.SQLPwd, "acedb", "System.Data.SqlClient");
            this.bisConnection = new HzConexao(bissql.SQLHost, bissql.SQLUser, bissql.SQLPwd, "BISEventLog", "System.Data.SqlClient");
            this.bisManager = bismanager;
        }
        #endregion    
    }
}