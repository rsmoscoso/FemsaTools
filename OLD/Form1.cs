using AMSLib;
using AMSLib.ACE;
using AMSLib.Common;
using AMSLib.Manager;
using AMSLib.SQL;
using HzBISCommands;
using HzLibConnection.Data;
using HzUtilClasses.Threading;
using Newtonsoft.Json;
using Serilog;
using SG3ServiceReference;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace FemsaTools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        #region Variables
        /// <summary>
        /// Log das tarefas.
        /// </summary>
        public ILogger LogTask { get; set; }
        /// <summary>
        /// Log das tarefas.
        /// </summary>
        public ILogger LogTaskSolo { get; set; }
        /// <summary>
        /// Resultado dos bloqueios/desbloqueios.
        /// </summary>
        public ILogger LogResult { get; set; }
        /// <summary>
        /// Resultado dos bloqueios/desbloqueios.
        /// </summary>
        public ILogger LogResultSolo { get; set; }
        /// <summary>
        /// Log das tarefas de importação.
        /// </summary>
        public ILogger LogTaskImport { get; set; }
        /// <summary>
        /// Resultado da importação.
        /// </summary>
        public ILogger LogResultImport { get; set; }
        /// <summary>
        /// Thread do bloqueio.
        /// </summary>
        private ObjectTrigger trigger { get; set; }
        /// <summary>
        /// Objeto da thread do horário para a execução da trigger de importação.
        /// </summary>
        private ObjectTrigger TrSchedule { get; set; }
        /// <summary>
        /// Configuração do servidor SQL Server Workforce.
        /// </summary>
        private SQLParameters workForceSQL { get; set; }
        /// <summary>
        /// Configuração do servidor SQL Server BISACE.
        /// </summary>
        private SQLParameters BISACESQL { get; set; }
        /// <summary>
        /// Configuração do servidor SQL Server BIS.
        /// </summary>
        private SQLParameters BISSQL { get; set; }
        /// <summary>
        /// Configuração do servidor do BIS.
        /// </summary>
        private BISManager BISManager { get; set; }
        /// <summary>
        /// Conexão com o banco de dados do BIS.
        /// </summary>
        public HzConexao bisACEConnection { get; set; }
        public HzConexao bisLOGConnection { get; set; }
        /// <summary>
        /// Dados para o compartilhamento do arquivo de importação.
        /// </summary>
        private FemsaNetworkCredential femsaCredential { get; set; }

        #endregion


        #region Functions
        /// <summary>
        /// Executa a leitura dos parâmetros do serviço.
        /// </summary>
        private void ReadParameters()
        {
            StreamReader reader = null;
            try
            {
                this.LogTask.Information("Lendo os parametros...");

                this.LogTask.Information("Lendo Parametros do BIS...");
                reader = new StreamReader(@"c:\horizon\config\bis.json");
                string sql = reader.ReadToEnd();
                reader.Close();

                if (String.IsNullOrEmpty(sql))
                    throw new Exception("Nao ha arquivo de configuracao do do BIS.");

                this.BISManager = JsonConvert.DeserializeObject<BISManager>(sql);
                this.LogTask.Information(String.Format("BIS --> Host: {0}, User: {1}, Pwd: {2}", this.BISManager.Host,
                    this.BISManager.User, this.BISManager.Pwd));

                this.LogTask.Information("Lendo SQL BISACE...");
                reader = new StreamReader(@"c:\horizon\config\sqlace.json");
                sql = reader.ReadToEnd();
                reader.Close();

                if (String.IsNullOrEmpty(sql))
                    throw new Exception("Nao ha arquivo de configuracao do SQL SERVER do BISACE.");

                this.BISACESQL = JsonConvert.DeserializeObject<SQLParameters>(sql);
                this.LogTask.Information(String.Format("SQL BIS --> Host: {0}, User: {1}, Pwd: {2}", this.BISACESQL.SQLHost,
                    this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd));

                this.LogTask.Information("Lendo SQL BIS...");
                reader = new StreamReader(@"c:\horizon\config\sql.json");
                sql = reader.ReadToEnd();
                reader.Close();

                if (String.IsNullOrEmpty(sql))
                    throw new Exception("Nao ha arquivo de configuracao do SQL SERVER do BIS.");

                this.BISSQL = JsonConvert.DeserializeObject<SQLParameters>(sql);
                this.LogTask.Information(String.Format("SQL BIS --> Host: {0}, User: {1}, Pwd: {2}", this.BISSQL.SQLHost,
                    this.BISSQL.SQLUser, this.BISSQL.SQLPwd));

                this.LogTask.Information("Lendo Arquivo de compartilhamento...");
                reader = new StreamReader(@"c:\horizon\config\credential.json");
                sql = reader.ReadToEnd();
                reader.Close();

                if (String.IsNullOrEmpty(sql))
                    throw new Exception("Nao ha arquivo com os dados do compartilhamento.");

                this.femsaCredential = JsonConvert.DeserializeObject<FemsaNetworkCredential>(sql);
                this.femsaCredential.SharingFolder = @"\\kofbrsmcappd00.sa.kof.ccf\Interfaces\RRHH";
                this.LogTask.Information(String.Format("Credential --> User: {0}, Pwd: {1}, Domain: {2}, SharingFolder: {3}, FileName: {4}", this.femsaCredential.UserName,
                    this.femsaCredential.Password, this.femsaCredential.Domain, this.femsaCredential.SharingFolder, this.femsaCredential.FileName));

                //this.LogTask.Information("Lendo Parametros do schedule...");
                //reader = new StreamReader(@"c:\horizon\config\schedule.json");
                //sql = reader.ReadToEnd();
                //reader.Close();

                //if (String.IsNullOrEmpty(sql))
                //    throw new Exception("Nao ha arquivo de configuracao do schedule.");

                //this.TimeSchedule = JsonConvert.DeserializeObject<Schedule>(sql);
                //this.LogTask.Information(String.Format("Schedule --> Data: {0}, Intervalo: {1}", this.TimeSchedule.StartTime,
                //    this.TimeSchedule.Interval.ToString()));
            }
            catch (Exception ex)
            {
                if (reader != null)
                    reader.Close();
                this.LogTask.Information(ex, ex.Message);
            }
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\BasePropriosFemsa.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\CheckCards_Proprios.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckCardsProprios.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");

                string line = null;
                string sql = null;
                bool bPersnoEqual = false;
                bool bCardnoEqual = false;
                bool bPersno = false;
                bool bCardno = false;
                string persid = null;
                reader.ReadLine();
                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}", "persid", "persno", "nome", "cardnoxls", "cardnobis", "persnoexists", "persnoequal", "cardexists", "cardequal", "cardid", "cardno", "sitecode"));
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persno = arr[0];
                    string nome = arr[1];
                    string cardno = arr[4].TrimStart(new Char[] { '0' });
                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Cardno: {2}", nome, persno, cardno));

                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select per.persid, persno, cardno, cardid from bsuser.persons per left outer join bsuser.cards cd on cd.persid = per.persid where (convert(bigint, cardno) = {0} or persno = '{1}') and cd.status = 1 and per.status = 1", cardno, persno)))
                    {
                        if (table.Rows.Count > 0)
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                bPersnoEqual = false;
                                bCardnoEqual = false;
                                if (bPersno = !String.IsNullOrEmpty(row["persno"].ToString()))
                                    bPersnoEqual = row["persno"].ToString().Equals(persno);
                                if (bCardno = !String.IsNullOrEmpty(row["cardid"].ToString()))
                                    bCardnoEqual = row["cardno"].ToString().TrimStart(new Char[] { '0' }).Equals(cardno);

                                using (DataTable tableBIS = this.bisLOGConnection.loadDataTable(String.Format("set dateformat 'dmy' exec spLastUsedCardID '{0}'", row["persid"].ToString())))
                                {
                                    if (tableBIS.Rows.Count > 0)
                                    {
                                        using (DataTable tableCard = this.bisACEConnection.loadDataTable(String.Format("select cardid, cardno, sitecode = convert(int,convert(varbinary(4), CODEDATA)) from bsuser.cards where cardid = '{0}'", tableBIS.Rows[0][0].ToString())))
                                        {
                                            this.LogTask.Information(String.Format("Registros encontrados {0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}", row["persid"].ToString(), persno, nome, cardno, row["cardno"].ToString().TrimStart(new Char[] { '0' }), bPersno, bPersnoEqual, bCardno, bCardnoEqual,
                                                table.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}", row["persid"].ToString(), persno, nome, cardno, row["cardno"].ToString().TrimStart(new Char[] { '0' }), bPersno, bPersnoEqual, bCardno, bCardnoEqual,
                                                tableBIS.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                        }
                                    }
                                    else
                                    {
                                        this.LogTask.Information("Cartao e persno nao encontrado.");
                                        writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}", "naoencontrado", persno, nome, cardno, false, false, false, false, false, false, false, false, false));
                                    }
                                }

                            }
                        }
                        else
                        {
                            this.LogTask.Information("Cartao e persno nao encontrado.");
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11}", "naoencontrado", persno, nome, cardno, false, false, false, false, false, false, false, false, false));
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\NotFoundSG3.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\CorrigidoSG3.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CleanSG3.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();


                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                CultureInfo ptBR = new CultureInfo("pt-BR", true);

                string line = null;
                string persno = "";
                string nome = null;
                string cpf = null;
                string rg = null;
                string cardno = null;
                string status = null;
                string id = null;
                string acesso = null;
                string sql = null;
                bool bdelete = false;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    if (arr.Length < 8)
                        continue;

                    status = arr[0];
                    id = arr[1];
                    persno = arr[2];
                    nome = arr[3];
                    cardno = arr[4];
                    acesso = arr[5];
                    cpf = arr[6];
                    rg = arr[7];
                    bdelete = false;
                    this.LogTask.Information(String.Format("Verificando Status: {0}, PERSID: {1}, Nome: {2}, Persno: {3}, CPF: {4}, RG: {5}, Cardno: {6}, Acesso: {7}", status, id, nome, persno, cpf, rg, cardno, acesso));

                    if (status.ToLower().Equals("naoencontrado"))
                        continue;

                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select persid, persclass, status from bsuser.persons where persid = '{0}'", id)))
                    {
                        if (table.Rows.Count > 0)
                        {
                            if (table.Rows[0]["persclass"].ToString().ToLower().Equals("v"))
                            {
                                this.LogTask.Information("Usuario e um visitante.");
                                continue;
                            }
                            else if (table.Rows[0]["status"].ToString().ToLower().Equals("0"))
                            {
                                this.LogTask.Information("Registro excluido.");
                                continue;
                            }
                        }
                        else
                            continue;
                    }

                    int days = -1;
                    days = acesso.ToLower().Equals("semultimoacesso") ? 60 : DateTime.Now.Subtract(DateTime.Parse(acesso, ptBR)).Days;

                    if (days >= 60 || cardno.ToLower().Equals("semcartao"))
                    {
                        this.LogTask.Information("Colaborador sem acesso ha mais de 60 dias ou sem cartao. Excluindo...");
                        BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                        this.LogTask.Information("Carregando o colaborador...");
                        if (per.Get(id) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                        {
                            this.LogTask.Information("Removendo todos os cartoes...");
                            if (per.RemoveAllCards() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                            {
                                this.LogTask.Information("Cartoes removidos com sucesso.");
                                this.LogTask.Information("Removendo o colaborador...");
                                if (per.Delete() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                                    this.LogTask.Information("Colaborador removido com sucesso.");
                                else
                                    this.LogTask.Information("Erro ao remover o colaborador.");
                            }
                        }

                        continue;
                    }

                    this.LogTask.Information("Verificando a existencia do persno.");
                    sql = String.Format("select persid, persno, persclassid from bsuser.persons where persno = '{0}'", persno);
                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table.Rows.Count > 0)
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                if (!row["persno"].ToString().Equals(persno))
                                {
                                    this.LogTask.Information(String.Format("Persno diferente. Persno: {0}. Verificando o cardno: {1}.", row["persno"].ToString(), cardno));
                                    sql = String.Format("select per.persid, persclassid, persno from bsuser.persons per inner join bsuser.cards cd on cd.persid = per.persid where convert(bigint, cardno) = {0}",
                                        long.Parse(cardno));
                                    using (DataTable tableCard = this.bisACEConnection.loadDataTable(sql))
                                    {
                                        if (tableCard.Rows.Count > 0)
                                        {
                                            this.LogTask.Information(String.Format("Cartao encontrado. Persno{0}, Persclassid: {1}", tableCard.Rows[0]["persno"].ToString(), tableCard.Rows[0]["persclassid"].ToString()));
                                        }
                                        else
                                        {
                                            this.LogTask.Information("Cartao nao encontrado.");
                                        }
                                    }
                                }
                                else
                                {
                                    this.LogTask.Information("Persno existente e igual. Persclassid: {0}", row["persclassid"].ToString());
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information("Rotina finalizada.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\checklastcard_proprios.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\CardID_Proprios.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckLastUsedCardProprios.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string line = null;
                string sql = null;
                reader.ReadLine();
                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", "persid", "re", "nome", "cardid", "cardno", "sitecode"));
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persid = arr[0];
                    string persno = arr[1];
                    string nome = arr[2];
                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, PERSID: {2}", nome, persno, persid));

                    if (persid.Equals("naoencontrado"))
                        continue;

                    using (DataTable table = this.bisLOGConnection.loadDataTable(String.Format("set dateformat 'dmy' exec spLastUsedCardID '{0}'", persid)))
                    {
                        if (table.Rows.Count > 0)
                        {
                            using (DataTable tableCard = this.bisACEConnection.loadDataTable(String.Format("select cardid, cardno, sitecode = convert(int,convert(varbinary(4), CODEDATA)) from bsuser.cards where cardid = '{0}'", table.Rows[0][0].ToString())))
                            {
                                this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5}", nome, persno, persid, table.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", nome, persno, persid, table.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                            }
                        }
                        else
                        {
                            this.LogTask.Information("PERSID nao encontrado.");
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", nome, persno, persid, "naoencontrado", "naoencontrado", "naoencontrado"));
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\BasePropriosFemsa.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\CheckDuplicados_Proprios.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckDuplicados.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string line = null;
                reader.ReadLine();
                writer.WriteLine(String.Format("{0};{1};{2}", "persno", "nome", "qtd"));
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persno = arr[0];
                    string nome = arr[1];
                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}", nome, persno));

                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select total = count(*) from bsuser.persons where persno = '{0}' and status = 1 and persclass = 'E'", persno)))
                    {
                        if (table.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("Registro encontrado Nome: {0}, RE: {1}; QTD; {2}", nome, persno, table.Rows[0]["total"].ToString()));
                            writer.WriteLine(String.Format("{0};{1};{2}", persno, nome, table.Rows[0]["total"].ToString()));
                        }
                        else
                        {
                            this.LogTask.Information("Cartao e persno nao encontrado.");
                            writer.WriteLine(String.Format("{0};{1};{2};", persno, nome, "naoencontrado"));
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        /// <summary>
        /// Retorna os dados relativos ao cardid desejado.
        /// </summary>
        /// <param name="cardid">ID do cartão.</param>
        /// <returns></returns>
        private BSCardsInfo getCards(string cardid)
        {
            List<BSCardsInfo> retval = null;
            try
            {
                this.LogTask.Information(String.Format("Verificando os dados do CARDID: {0}", cardid));
                using (DataTable tableCard = this.bisACEConnection.loadDataTable(String.Format("select cardid, cardno, sitecode = convert(varchar, convert(int,convert(varbinary(4), CODEDATA))) from bsuser.cards where cardid = '{0}'", cardid)))
                {
                    if (tableCard.Rows.Count > 0)
                    {
                        retval = (List<BSCardsInfo>)BSCommon.ConvertDataTable<BSCardsInfo>(tableCard);
                        this.LogTask.Information(String.Format("CARDID encontrado. CARDNO: {0}, SITECODE: {1}", retval[0].CARDNO, retval[0].SITECODE));
                    }
                    else
                        this.LogTask.Information("CARDID nao encontrado.");
                }

                return retval[0];
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return null;
            }
        }

        /// <summary>
        /// Retorna o ID do cartão da pessoa.
        /// </summary>
        /// <param name="persid">ID do cartão.</param>
        /// <returns></returns>
        private string getCardID(string persid)
        {
            string retval = null;
            try
            {
                this.LogTask.Information(String.Format("Verificando o cardid da pessoa PERSID: {0}", persid));
                using (DataTable tableCard = this.bisACEConnection.loadDataTable(String.Format("select cardid from bsuser.cards where persid = '{0}'", persid)))
                {
                    if (tableCard.Rows.Count > 0)
                    {
                        retval = tableCard.Rows[0]["cardid"].ToString();
                        this.LogTask.Information(String.Format("CARDID encontrado: {0}", retval));
                    }
                    else
                        this.LogTask.Information("CARDID nao encontrado.");
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
        /// Retorna a data do último acesso da pessoa com o cartão 
        /// específico.
        /// </summary>
        /// <param name="persid">ID da pessoa.</param>
        /// <param name="cardid">ID do cartão.</param>
        /// <returns></returns>
        private DateTime? getLastAccess(string persid, string cardid)
        {
            DateTime? retval = null;
            try
            {
                this.LogTask.Information(String.Format("Verificando o ultimo acesso da pessoa com o cartao atual. PERSID: {0}, CARDID: {1}", persid, cardid));
                using (DataTable table = this.bisLOGConnection.loadDataTable(String.Format("set dateformat 'dmy' exec spLastUsedCardID_PERSID '{0}'", persid)))
                {
                    if (table.Rows.Count > 0)
                    {
                        retval = DateTime.Parse(table.Rows[0]["UltimoAcesso"].ToString());
                        this.LogTask.Information(String.Format("Data encontrada: {0}", retval.Value.ToString("dd/MM/yyyy HH:mm:ss")));
                    }
                    else
                        this.LogTask.Information("Nao ha data de acesso.");
                }

                return retval;
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return retval;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\Femsaterceiros.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\REDuplicados_CardIDResult.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\REDuplicadosCardID.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");

                string line = null;
                reader.ReadLine();
                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", "persno", "nome", "persid", "cardid", "cardno", "codedata", "ultimoacesso"));
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persno = arr[0];
                    string nome = arr[1];
                    string cardid = null;
                    //int qtd = int.Parse(arr[2].ToString());
                    int qtd = 0;
                    string persid = null;
                    DateTime? lastaccess = null;
                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Qtd.: {2}", nome, persno, qtd.ToString()));
                    using (DataTable tableID = this.bisACEConnection.loadDataTable(String.Format("select per.persid, cardid from bsuser.persons per left outer join bsuser.cards cd on cd.persid = per.persid " +
                        "where per.status = 1 and (CARDID is null or (not cardid is null and cd.status = 1)) and persno = '{0}'", persno)))
                    {
                        if (tableID.Rows.Count > 0)
                        {
                            int i = 1;
                            foreach (DataRow row in tableID.Rows)
                            {
                                this.LogTask.Information(String.Format("Verificando {0} de {1}", (i++).ToString(), tableID.Rows.Count.ToString()));
                                persid = row["persid"].ToString();
                                cardid = String.IsNullOrEmpty(row["cardid"].ToString()) ? null : row["cardid"].ToString();

                                if (cardid == null)
                                {
                                    this.LogTask.Information("CARDID nao encontrado.");
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", persno, nome, persid, "naoencontrado", "naoencontrado", "naoencontrado", "naoencontrado"));
                                    continue;
                                }

                                lastaccess = this.getLastAccess(persid, cardid);
                                BSCardsInfo card = this.getCards(cardid);

                                this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5};{6}", persno, nome, persid, cardid, card != null ? card.CARDNO : "naoencontrado",
                                    card != null ? card.SITECODE : "naoencontrado", lastaccess != null ? lastaccess.Value.ToString("dd/MM/yyyy HH:mm:ss") : "naoencontrado"));
                                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", persno, nome, persid, cardid, card != null ? card.CARDNO : "naoencontrado",
                                    card != null ? card.SITECODE : "naoencontrado", lastaccess != null ? lastaccess.Value.ToString("dd/MM/yyyy HH:mm:ss") : "naoencontrado"));
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\REDuplicados_CardIDResult.csv");
            List<RECard> recards = new List<RECard>();
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\RemoveDuplicados.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Conectando com o BIS.");
                if (this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string line = null;
                string persno = "";
                string nome;
                string persid;
                string cardid;
                string cardno;
                string sitecode;
                string acesso;
                reader.ReadLine();
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    persno = arr[0];
                    nome = arr[1];
                    persid = arr[2];
                    cardid = arr[3];
                    cardno = arr[4];
                    sitecode = arr[5];
                    acesso = arr[6];
                    this.LogTask.Information(String.Format("Criando o dicionario Persid: {0}, Nome: {1}, RE: {2}, Cardid: {3}, Cardno: {4}, Sitecode: {5}; Acesso: {6}", persid, nome, persno, cardid, cardno, sitecode, acesso));

                    if (acesso.Equals("naoencontrado"))
                        continue;

                    List<RECard> result = recards.FindAll(delegate (RECard c) { return c.Persno.Equals(persno); });
                    DateTime maxAcesso = DateTime.Now;
                    string maxCardid = "";
                    string maxPersid = "";
                    if (result != null && result.Count > 0)
                    {
                        maxCardid = "";
                        maxPersid = "";
                        foreach (RECard cd in result)
                        {
                            if (cd.Acesso.Subtract(DateTime.Parse(acesso)).Hours > 0)
                            {
                                maxAcesso = DateTime.Parse(acesso);
                                maxCardid = cardid;
                                maxPersid = persid;
                            }
                        }

                        recards.Find(delegate (RECard c) { return c.Persno.Equals(persno); }).Acesso = maxAcesso;
                        recards.Find(delegate (RECard c) { return c.Persno.Equals(persno); }).Cardid = maxCardid;
                        recards.Find(delegate (RECard c) { return c.Persno.Equals(persno); }).Persid = maxPersid;
                    }
                    else
                        recards.Add(new RECard() { Persno = persno, Cardid = cardid, Acesso = DateTime.Parse(acesso) });
                }

                this.LogTask.Information("Iniciando a exclusao dos duplicados.");
                foreach (RECard cd in recards)
                {
                    BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                    this.LogTask.Information(String.Format("Excluindo Persid: {0}, Persno: {1}, Cardid: {2}, Acesso: {3}", cd.Persid, cd.Persno, cd.Cardid, cd.Acesso.ToString("dd/MM/yyyy HH:mm:ss")));
                    this.LogTask.Information("Carregando a pessoa.");
                    if (per.Load(cd.Persid))
                    {
                        this.LogTask.Information("Excluindo a pessoa.");
                        if (per.DeletePerson(cd.Persid) == BS_ADD.SUCCESS)
                            this.LogTask.Information("Pessoa excluida com sucesso.");
                        else
                            this.LogTask.Information(String.Format("Erro ao excluir a pessoa. {0}", per.GetErrorMessage()));
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                ////this.LogTask.Information("Descontecando o BIS.");
                ////if (this.BISManager.Disconnect())
                ////    this.LogTask.Information("BIS desconectado com sucesso.");
                ////else
                ////    this.LogTask.Information("Erro ao desconectar com o BIS.");
                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\REDuplicados_CardIDResult.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\RemoveDuplicadosNotFound.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string line = null;
                string persno;
                string nome;
                string persid;
                string cardid;
                string cardno;
                string sitecode;
                string acesso;
                reader.ReadLine();
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    persno = arr[0];
                    nome = arr[1];
                    persid = arr[2];
                    cardid = arr[3];
                    cardno = arr[4];
                    sitecode = arr[5];
                    acesso = arr[6];
                    this.LogTask.Information(String.Format("Verificando Persid: {0}, Nome: {1}, RE: {2}, Cardid: {3}, Cardno: {4}, Sitecode: {5}; Acesso: {6}", persid, nome, persno, cardid, cardno, sitecode, acesso));

                    if (cardid.Equals("naoencontrado"))
                    {
                        if (String.IsNullOrEmpty(this.getCardID(persid)))
                        {
                            BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                            this.LogTask.Information("Carregando a pessoa.");
                            if (per.Load(persid))
                            {
                                this.LogTask.Information("Pessoa carregada com sucesso.");
                                this.LogTask.Information("Excluindo a pessoa.");
                                if (per.Delete() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                                    this.LogTask.Information("Pessoa excluida com sucesso.");
                                else
                                    this.LogTask.Information("Erro ao exlcuir a pessoa.");
                            }
                            else
                                this.LogTask.Information("Erro ao carregar a pessoa.");
                        }
                    }
                    else
                        this.LogTask.Information("Pessoa com cartao valido.");
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                reader.Close();
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
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
                using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where codigo = '{0}'", codigo.TrimStart(new Char[] { '0' }))))
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
        private void button9_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\BasePropriosFemsa.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\CheckDivisao.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckDivisao.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");

                string line = null;
                reader.ReadLine();
                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", "persid", "persno", "nome", "Area", "Descricao", "Divisao BIS", "Comparacao"));
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persno = arr[0];
                    string nome = arr[1];
                    string area = arr[9];
                    string descricao = arr[10];
                    string clientid = null;

                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Area: {2}, Descricao: {3}", nome, persno, area, descricao));

                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select per.persid, per.clientid, cli.name from bsuser.persons per inner join bsuser.clients cli on cli.clientid = per.clientid where persno = '{0}' and per.status = 1", persno)))
                    {
                        if (table.Rows.Count > 0)
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                if ((clientid = this.getClientidFromArea(area)) != null)
                                {
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", row["persid"].ToString(), persno, nome, area, descricao, row["name"].ToString(),
                                        clientid.ToLower().Trim().Equals(row["clientid"].ToString().ToLower().Trim()) ? "S" : "N"));
                                    this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5};{6}", row["persid"].ToString(), persno, nome, area, descricao, row["name"].ToString(),
                                        clientid.ToLower().Trim().Equals(row["clientid"].ToString().ToLower().Trim()) ? "S" : "N"));
                                }
                                else
                                {
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", row["persid"].ToString(), persno, nome, area, descricao, "naoencontrado", "N"));
                                }
                            }
                        }
                        else
                        {
                            this.LogTask.Information("Persno nao encontrado.");
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", "naoencontrado", persno, nome, area, descricao, "naoencontrado", "N"));
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\AccessNotFound.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\AccessNotFound_Erro.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\RemoveAccessNotFound.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");

                //this.LogTask.Information("Conectando com o BIS.");
                //if (!this.BISManager.Connect())
                //    throw new Exception("Erro ao conectar com o BIS.");

                //this.LogTask.Information("BIS conectado com sucesso!");

                string line = null;
                string persno;
                string nome;
                string persid;
                string cardid;
                string cardno;
                string sitecode;
                string acesso;
                DateTime? lastAccess = null;
                reader.ReadLine();
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    persno = arr[0];
                    nome = arr[1];
                    persid = arr[2];
                    cardid = arr[3];
                    cardno = arr[4];
                    sitecode = arr[5];
                    acesso = arr[6];
                    this.LogTask.Information(String.Format("Verificando Persid: {0}, Nome: {1}, RE: {2}, Cardid: {3}, Cardno: {4}, Sitecode: {5}; Acesso: {6}", persid, nome, persno, cardid, cardno, sitecode, acesso));

                    if (persno.Equals("453858"))
                        acesso = acesso;
                    if (!String.IsNullOrEmpty(this.getCardID(persid)))
                    {
                        this.LogTask.Information("Verificando o ultimo acesso.");
                        lastAccess = this.getLastAccess(persid, cardid);
                        if (lastAccess != null && lastAccess.Value.Year < 2022)
                        {
                            this.LogTask.Information(String.Format("Ultimo acesso nao encontrado ou inferior a 2022: {0}", lastAccess == null ? "Nao houve acesso." : lastAccess.Value.ToString("dd/MM/yyyy HH:mm:ss")));
                            //BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                            //this.LogTask.Information("Carregando a pessoa.");
                            //if (per.Load(persid))
                            //{
                            //    this.LogTask.Information("Pessoa carregada com sucesso.");
                            //    this.LogTask.Information("Excluindo a pessoa.");
                            //    if (per.Delete() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                            //        this.LogTask.Information("Pessoa excluida com sucesso.");
                            //    else
                            //        this.LogTask.Information("Erro ao exlcuir a pessoa.");
                            //}
                            //else
                            //    this.LogTask.Information("Erro ao carregar a pessoa.");
                        }
                        else
                        {
                            if (acesso.Equals("naoencontrado"))
                                writer.WriteLine(String.Format("Verificando Persid: {0}, Nome: {1}, RE: {2}, Cardid: {3}, Cardno: {4}, Sitecode: {5}; Acesso: {6}", persid, nome, persno, cardid, cardno, sitecode, acesso));
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                reader.Close();
                writer.Close();
                //this.LogTask.Information("Descontecando o BIS.");
                //if (this.BISManager.Disconnect())
                //    this.LogTask.Information("BIS desconectado com sucesso.");
                //else
                //    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //StreamReader reader = new StreamReader(@"c:\temp\CheckDivisao.csv");
            StreamReader reader = new StreamReader(@"c:\temp\FemsaProprios_Terc.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\CheckDivisaoMoreThan200.csv");
            //StreamWriter writerC = new StreamWriter(@"c:\temp\CheckDivisaoAuths_Count.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\ChangeDivision.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string line = null;
                reader.ReadLine();
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    //string persid = arr[0];
                    string persno = arr[0];
                    string nome = arr[1];
                    //string area = arr[3];
                    string descricao = arr[4];
                    //string divisaobis = arr[5];
                    //string comparacao = arr[6];
                    //string divisaocorreta = arr[7];
                    string clientid = null;
                    string authstr = null;

                    //this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Area: {2}, Descricao: {3}, Divisao BIS: {4}", nome, persno, area, descricao, divisaobis));
                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Descricao: {2}", nome, persno, descricao));

                    //this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Area: {2}, Descricao: {3}, Divisao BIS: {4}, Divisao Correta: {5}", nome, persno, area, descricao, divisaobis, divisaocorreta));

                    //if (divisaobis.ToLower().Equals(divisaocorreta.ToLower()))
                    //{
                    //    this.LogTask.Information("Divisao correta.");
                    //    continue;
                    //}

                    //using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select persid from bsuser.persons where persno = '{0}' and status = 1", persno)))
                    //{
                    //    foreach (DataRow row in table.Rows)
                    //    {
                    //        List<BSAuthorizationInfo> auths = BSSQLPersons.GetAllAuthorizations(this.bisACEConnection, row["persid"].ToString());
                    //        if (auths == null || auths.Count < 1)
                    //        {
                    //            writer.WriteLine(String.Format("{0},{1}", row["persid"].ToString(), persno));
                    //            writer.WriteLine(String.Format("{0},{1};{2}", row["persid"].ToString(), persno, "0"));
                    //        }
                    //        else
                    //        {
                    //            authstr = String.Format("{0},{1},", row["persid"].ToString(), persno);
                    //            foreach (BSAuthorizationInfo auth in auths)
                    //            {
                    //                authstr += String.Format("{0},", auth.AUTHID);
                    //            }
                    //            writer.WriteLine(authstr.Substring(0, authstr.Length - 1));
                    //            writerC.WriteLine(String.Format("{0},{1},{2}", row["persid"].ToString(), persno, auths.Count.ToString()));
                    //        }
                    //    }
                    //}
                    //continue;

                    //this.LogTask.Information("Verificando o ID do cliente.");
                    //clientid = BSSQLClients.GetID(this.bisACEConnection, divisaocorreta);

                    //if (String.IsNullOrEmpty(clientid))
                    //{
                    //    this.LogTask.Information("Divisao nao encontrada.");
                    //    continue;
                    //}

                    this.LogTask.Information("Verificando o ID do cliente.");
                    if (descricao.ToLower().Equals("planta jundiaí"))
                        descricao = "PLANTA JUNDIAI";

                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where area = '{0}'", descricao)))
                    {
                        if (table.Rows.Count > 0)
                            clientid = table.Rows[0]["clientid"].ToString();
                        else
                        {
                            if (String.IsNullOrEmpty(clientid))
                            {
                                this.LogTask.Information("Divisao nao encontrada.");
                                continue;
                            }

                        }
                    }

                    clientid = "0012FC697204689A";
                    string persid = BSSQLPersons.GetPersid(this.bisACEConnection, "5157775");
                    List<BSAuthorizationInfo> auths = BSSQLPersons.GetAllAuthorizations(this.bisACEConnection, persid);
                    int totalBefore = auths != null ? auths.Count : 0;
                    int totalAfter = -1;

                    if (totalBefore > 200)
                    {
                        this.LogTask.Information(String.Format("Colaborador com mais de 200 autorizacoes. Qtd: {0}", totalBefore.ToString()));
                        //writer.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6}", nome, persno, area, descricao, divisaobis, divisaocorreta, totalBefore.ToString()));
                        continue;
                    }

                    BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                    this.LogTask.Information("Apontando a pessoa.");
                    if (persons.Get(persid) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                    {
                        // VERIFICAR QTD DE AUTORIZACOES. MAXIMO 20
                        this.LogTask.Information("Alterando a divisao.");
                        if (persons.ChangeClient(clientid) == BS_SAVE.SUCCESS)
                        {
                            this.LogTask.Information("Divisao alterada com sucesso.");
                            auths = BSSQLPersons.GetAllAuthorizations(this.bisACEConnection, persid);
                            totalAfter = auths != null ? auths.Count : 0;

                            if (totalAfter != totalBefore)
                                throw new Exception(String.Format("Total Antes: {0}, Total Depois: {1}. Erro na alteracao. Rotina interrompida.", totalBefore.ToString(), totalAfter.ToString()));
                        }
                        else
                            this.LogTask.Information("Erro ao alterar a divisao.");
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();
                //writerC.Close();
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //StreamReader reader = new StreamReader(@"c:\temp\Bosch_16.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\ImportProprios.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();

                StreamReader reader = new StreamReader(@"c:\temp\Bosch_3");
                string line = null;
                BSPersonsInfoFEMSA personsFemsa = null;
                List<BSPersonsInfoFEMSA> list = new List<BSPersonsInfoFEMSA>();
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

                    list.Add(personsFemsa);
                }
                reader.Close();
                reader = null;
                FEMSAImport femsa = new FEMSAImport(this.LogTask, this.LogResult, this.BISACESQL, this.BISSQL, this.BISManager);
                //                femsa.Import(this.femsaCredential);
                femsa.ImportList(list);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                //reader.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AddAuthorizaion.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Lendo o arquivo com as autorizacoes.");
                List<BSPersonsInfo> personsInfo = Global.readCSV(openFileDialog1);
                this.LogTask.Information(String.Format("{0} autorizacao(oes) encontrada(s)", personsInfo == null || personsInfo.Count < 1 ? "0" : personsInfo.Count.ToString()));

                this.Cursor = Cursors.WaitCursor;

                if (personsInfo == null || personsInfo.Count < 1)
                    throw new Exception("Nao ha autorizacoes para adicionar.");

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                foreach (BSPersonsInfo info in personsInfo)
                {
                    BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                    this.LogTask.Information(String.Format("Carregando o colaborador: Persno: {0}", info.PERSNO));
                    if (persons.Load(info.PERSNO, BS_STATUS.ACTIVE))
                    {
                        this.LogTask.Information("Colaborador carregado com sucesso.");
                        foreach (BSAuthorizationInfo authinfo in info.AUTHORIZATIONS)
                        {
                            this.LogTask.Information(String.Format("Verificando a autorizacao: Nome: {0}", authinfo.SHORTNAME));
                            string authid = BSSQLAuthorizations.GetID(this.bisACEConnection, authinfo.SHORTNAME);
                            if (String.IsNullOrEmpty(authid))
                            {
                                this.LogTask.Information("Autorizacao nao encontrada.");
                                break;
                            }

                            this.LogTask.Information(String.Format("Adicionando a autorizacao: ID: {0}, Nome: {1}", authid, authinfo.SHORTNAME));
                            if (persons.AppendAuthorization(authid) == BS_ADD.SUCCESS)
                                this.LogTask.Information("Autorizacao adicionada com sucesso.!");
                            else
                                this.LogTask.Information("Erro ao adicionar a autorizacao.");
                        }
                    }
                    else
                        this.LogTask.Information("Erro ao carregar o colaborador.");
                }
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        public DataTable getAllAuthPersno(string persno)
        {
            try
            {
                return this.bisLOGConnection.loadDataTable(String.Format("set dateformat 'dmy' exec spGetAuthsPersno '{0}'", persno));
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                return null;
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AdjustDivisaoBH.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");

                this.Cursor = Cursors.WaitCursor;
                string line = null;
                string persid = null;
                string persno = null;
                List<string> auths = null;
                string clientid = null;
                string authid = "0013B644E3CC1D3E";
                string[] files = Directory.GetFiles(@"C:\temp\RONALDO\jur geral");
                BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                int bytesCount = 0;
                foreach(string file in files)
                {
                    FileInfo f = new FileInfo(file);
                    reader = new StreamReader(file);
                    reader.ReadLine();
                    this.LogTask.Information(String.Format("Arquivo: {0}, Tamanho: {1}", file, f.Length.ToString()));
                    bytesCount = 0;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] arr = line.Split(',');
                        persno = arr[0];
                        bytesCount += System.Text.Encoding.Unicode.GetBytes(line).Length;
                        this.LogTask.Information(String.Format("{0} de {1}", bytesCount.ToString(), f.Length.ToString()));
                        this.LogTask.Information(String.Format("Verificando a autorizacao da pessoa: {0}", persno));
                        if (!String.IsNullOrEmpty(arr[0]))
                        {
                            this.LogTask.Information("Verificando a existencia da autorizacao.");
                            using (DataTable tbl = this.bisACEConnection.loadDataTable(String.Format("select per.persid from bsuser.authperperson aper " +
                                "inner join bsuser.persons per on per.persid = aper.persid where authid = '{0}' and per.persno = '{1}'", authid, persno)))
                            {
                                if (tbl == null || tbl.Rows.Count < 1)
                                {
                                    this.LogTask.Information("Autorizacao inexistente.");
                                    this.LogTask.Information("Carregando a pessoa");
                                    if (per.Load(arr[0], BS_STATUS.ACTIVE))
                                    {
                                        this.LogTask.Information("Pessoa carregada com sucesso.");
                                        if (per.AddAuthorization(authid) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                                            this.LogTask.Information("Autorizacao atribuida com sucesso.");
                                        else
                                            this.LogTask.Information("Erro ao atribuir a autorizacao.");
                                    }
                                }
                                else
                                    this.LogTask.Information("Autorizacao ja atribuida.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\ImportacaoFemsa.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\ImportacaoFemsa_ErroCheck.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\ErrosImportacaoFemsa.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string line = null;
                writer.WriteLine("Persno SAP, Nome SAP, CardNO SAP, Persid BIS, Persno BIS, Nome BIS");
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persno = arr[0];
                    string nome = arr[1];
                    string cardno = arr[2];
                    string status = arr[3];

                    this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, Cardno: {2}, Status: {3}", nome, persno, cardno, status));

                    if (status.ToLower().Equals("ok"))
                    {
                        this.LogTask.Information("Status ok.");
                        continue;
                    }

                    if (String.IsNullOrEmpty(cardno))
                    {
                        writer.WriteLine(String.Format("{0},{1},{2},{3},{4},{5}", persno, nome, "SEMCARTAO", "", "", ""));
                        continue;
                    }

                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select per.persid, persno, Nome = firstname + ' ' + lastname from bsuser.cards cd " +
                        "inner join bsuser.persons per on per.persid = cd.persid where convert(bigint, cardno) = {0}", cardno)))
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            writer.WriteLine(String.Format("{0},{1},{2},{3},{4},{5}", persno, nome, cardno, row["persid"].ToString(), row["persno"].ToString(), row["nome"].ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
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
                using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where area = '{0}'", area)))
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

        private async void button15_ClickAsync(object sender, EventArgs e)
        {
            StreamWriter writer = new StreamWriter(@"c:\temp\liberacao.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3Check.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();


                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                CultureInfo ptBR = new CultureInfo("pt-BR", true);

                this.LogTask.Information("Verificando o SG3.");
                //apenas crie o client
                var client = new ExecutivaPortClient();
                //e chame este método assíncrono com usuario, senha e o numero 3
                var response = await client.getLiberacaoWithCredentials(@"webservice.femsabrasil", @"T8yP@2jK$4r", 3);
                var x = response.FindAll(delegate (Liberacao l) { return l.liberado.Equals("S") && l.status_alocacao.Equals("S"); });

                string sql = null;
                string firstname = "";
                string lastname = "";
                writer.WriteLine("RG SG3;CPF SG3;Nome SG3;LIBERADO;PERSID;PERSNO;NOME;PERSCLASSID;ID");
                int i = 0;
                foreach (Liberacao l in x)
                {
                    ++i;

                    BSCommon.ParseName(l.colaborador.Replace("'", ""), out firstname, out lastname);

                    this.LogTask.Information(String.Format("{0} de {1}. RG: {2}, CPF: {3}, Nome: {4}", (i++).ToString(), x.Count.ToString(), l.rg, l.cpf, l.colaborador.Replace("'", "")));

                    sql = String.Format("select persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), persclassid from bsuser.persons where status = 1 and persclass = 'E' and (persno = '{0}' or persno = '{1}')",
                        l.rg.Replace("'", ""), l.cpf.Replace("'", ""), firstname, lastname);
                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                            foreach (DataRow r in table.Rows)
                            {
                                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7}", l.rg, l.cpf, l.colaborador, l.liberado, r["persid"].ToString(), r["persno"].ToString(), r["nome"].ToString(), r["persclassid"].ToString(), l.id_colaborador));
                            }
                        }
                        else
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7}", l.rg, l.cpf, l.colaborador, l.liberado, "sempersid", "sempersno", "semnome", "sempersclassid", l.id_colaborador));
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            string line = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AddAuthorizaion.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Lendo o arquivo com as autorizacoes.");
                List<BSPersonsInfo> personsInfo = Global.readCSV(openFileDialog1);

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                foreach (BSPersonsInfo info in personsInfo)
                {
                    BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                    this.LogTask.Information(String.Format("Carregando o colaborador: Persno: {0}", info.PERSNO));
                    if (persons.Load(info.PERSNO, BS_STATUS.ACTIVE))
                    {
                        this.LogTask.Information("Colaborador carregado com sucesso.");
                        foreach (BSAuthorizationInfo authinfo in info.AUTHORIZATIONS)
                        {
                            this.LogTask.Information(String.Format("Verificando a autorizacao: Nome: {0}", authinfo.SHORTNAME));
                            string authid = BSSQLAuthorizations.GetID(this.bisACEConnection, authinfo.SHORTNAME);
                            if (String.IsNullOrEmpty(authid))
                            {
                                this.LogTask.Information("Autorizacao nao encontrada.");
                                break;
                            }

                            if (persons.RemoveAuthorization(authid) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                                this.LogTask.Information("Autorizacao removida com sucesso.");
                            else
                                this.LogTask.Information("Erro ao remover a autorizacao.");
                        }
                    }
                    else
                        this.LogTask.Information("Erro ao carregar o colaborador.");
                }
                //reader = new StreamReader(openFileDialog1.FileName);
                //while ((line = reader.ReadLine()) != null)
                //{
                //    string sql = String.Format("select per.persid, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), aper.authid from bsuser.persons per inner join bsuser.authperperson aper on aper.persid = per.persid inner join bsuser.authorizations auth on auth.authid = aper.authid where shortname = '{0}'",
                //        line);
                //    this.LogTask.Information(String.Format("Verificando a existenica da autorizacao: {0}", line));
                //    BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                //    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                //    {
                //        if (table.Rows.Count > 0)
                //        {
                //            int i = 1;
                //            this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                //            foreach(DataRow row in table.Rows)
                //            {
                //                this.LogTask.Information(String.Format("{0} de {1}", (i++).ToString(), table.Rows.Count.ToString()));
                //                this.LogTask.Information(String.Format("Colaborador Persid:{0}, Nome: {1}, Authid: {2}", row["persid"].ToString(), row["nome"].ToString(), row["authid"].ToString()));
                //                this.LogTask.Information("Carregando o colaborador.");
                //                if (persons.Load(row["persid"].ToString()))
                //                {
                //                    this.LogTask.Information("Colaborador carregado com sucesso.");
                //                    if (persons.RemoveAuthorization(row["authid"].ToString()) == API_RETURN_CODES_CS.API_SUCCESS_CS)
                //                        this.LogTask.Information("Autorizacao removida com sucesso.");
                //                    else
                //                        this.LogTask.Information("Erro ao remover a autorizacao.");
                //                }
                //                else
                //                    this.LogTask.Information("Erro ao carregar o colaborador.");
                //            }
                //        }
                //        else
                //            this.LogTask.Information("Nenhuma pessoa encontrada com essa autorizacao.");
                //    }
                //}

                this.Cursor = Cursors.WaitCursor;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            string line = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\BlockPeople.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Lendo o arquivo com os bloqueios.");
                reader = new StreamReader(@"c:\temp\femsabloqueio.csv");

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string sql = null;
                string firstname = "";
                string lastname = "";

                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');
                    string nome = arr[0];
                    string cpf = arr[1].Replace(".", "").Replace("/", "").Replace("-", "");
                    string rg = arr[2];
                    this.LogTask.Information(String.Format("RG: {0}, CPF: {1}, Nome: {2}", rg, cpf, nome));

                    BSCommon.ParseName(nome, out firstname, out lastname);

                    sql = String.Format("select persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, '') from bsuser.persons where status = 1 and persclass = 'E' and ((persno = '{0}' or persno = '{1}') or (firstname like '%{2}%' and lastname like '%{3}%'))",
                        rg, cpf, firstname, lastname);
                    BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("Carregando o colaborador: Persid: {0}", table.Rows[0]["persid"].ToString()));
                            if (persons.Load(table.Rows[0]["persid"].ToString()))
                            {
                                this.LogTask.Information("Bloqueando a pessoa.");
                                if (persons.AddBlock(DateTime.Now, DateTime.Now.AddYears(10), "Documentacao Incompleta.") == BS_ADD.SUCCESS)
                                    this.LogTask.Information("Pessoa bloqueada com sucesso.");
                                else
                                    this.LogTask.Information("Erro bloqueando a pessoa.");
                            }
                        }
                    }
                }

                this.Cursor = Cursors.WaitCursor;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\ImportSG3Excel.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();

                FEMSAImport femsa = new FEMSAImport(this.LogTask, this.LogResult, this.BISACESQL, this.BISSQL, this.BISManager);
                femsa.ImportSG3Excel();
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                //reader.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\Bosch_27");
            StreamWriter writer = new StreamWriter(@"c:\temp\CheckProprios_Erro2601_Exato.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckPropriosCards.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string sql = null;
                string firstname = "";
                string lastname = "";
                int i = 1;
                string line = null;
                string clientid = null;
                BSPersonsInfoFEMSA personsFemsa = new BSPersonsInfoFEMSA();
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\Bosch_27");
                writer.WriteLine("PERSID,NOME BIS,RE BIS,CARDNO BIS,NOME SAP,RE SAP,CARDNO SAP,DIVISAO");

                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split('|');

                    string persno = arr[0].TrimStart(new Char[] { '0' });
                    string nome = arr[1].Replace("'", "");
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
                        NOME = nome.Replace("'", ""),
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

                    BSCommon.ParseName(nome, out firstname, out lastname);

                    this.LogTask.Information(String.Format("{0} de {1}. Nome: {2}, Cardno: {3}, RE: {4}, Status: {5}", (i++).ToString(), totallines.ToString(), nome, !String.IsNullOrEmpty(cardno) ? cardno : "SEMCARTAO", persno, status));

                    if (status.Equals("saiu da empresa"))
                        continue;

                    this.LogTask.Information(String.Format("Verificando o clientid da area: {0}", personsFemsa.CODIGOAREA.TrimStart(new Char[] { '0' })));
                    using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select clientid from Horizon..tblAreas where codigo = '{0}'", personsFemsa.CODIGOAREA.TrimStart(new Char[] { '0' }))))
                    {
                        if (table.Rows.Count > 0)
                        {
                            clientid = table.Rows[0]["clientid"].ToString();
                            this.LogTask.Information(String.Format("Area encontrada: {0}", clientid));
                        }
                        else
                            this.LogTask.Information("Area nao encontrada.");
                    }

                    sql = String.Format("select per.persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), cardno, cli.name from bsuser.persons per left outer join bsuser.cards cd on per.persid = cd.persid " +
                        " inner join bsuser.clients cli on cli.clientid = per.clientid " +
                        " where per.status = 1 and persclass = 'E' and persclassid = 'FF0000E500000001' and companyid in ('0013184849D0FE94', '00134B2266C87B31') and cd.status = 1 and ((firstname = '{0}' and lastname = '{1}') or (convert(bigint, cardno) = '{2}'))",
                        firstname, lastname, cardno);
                    if (!String.IsNullOrEmpty(clientid))
                        sql += string.Format(" and per.clientid = '{0}'", clientid);

                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table.Rows.Count > 0)
                        {
                            foreach (DataRow r in table.Rows)
                            {
                                writer.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6},{7}", r["persid"].ToString(), r["nome"].ToString(), r["persno"].ToString(), !String.IsNullOrEmpty(r["cardno"].ToString()) ? r["cardno"].ToString() : "SEMCARTAO", personsFemsa.NOME, personsFemsa.PERSNO,
                                    !String.IsNullOrEmpty(personsFemsa.CARDNO) ? personsFemsa.CARDNO : "SEMCARTAO", r["name"].ToString()));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\changeRE.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\changeRE_Erros.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\changeRE.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string sql = null;
                string firstname = "";
                string lastname = "";
                int i = 1;
                string line = null;
                BSPersonsInfoFEMSA personsFemsa = new BSPersonsInfoFEMSA();
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\changeRE.csv");

                writer.WriteLine("RE BIS, RE SAP, Nome");
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persid = arr[0];
                    string nomeBIS = arr[1];
                    string persnoBIS = arr[2];
                    string cardnoBIS = arr[3];
                    string nomeSAP = arr[4];
                    string persnoSAP = arr[5];
                    string cardnoSAP = arr[6];

                    this.LogTask.Information(String.Format("{0} de {1}. Persid: {2}, Nome: {3}, Cardno: {4}, RE BIS: {5}, RE SAP: {6}", (i++).ToString(), totallines.ToString(), persid, nomeBIS, cardnoSAP, persnoBIS, persnoSAP));

                    sql = String.Format("update bsuser.persons set persno = '{0}' where persid = '{1}'", persnoSAP, persid);
                    if (this.bisACEConnection.executeProcedure(sql))
                        this.LogTask.Information("Registro atualizado com sucesso.");
                    else
                    {
                        this.LogTask.Information("Erro ao atualizar o registro.");
                        writer.WriteLine(String.Format("Persid: {0}, RE BIS: {1}, RE SAP: {2}, Nome: {3}", persid, persnoBIS, persnoSAP, nomeBIS));
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        public string ChangeString(string str)
        {
            /** Troca os caracteres acentuados por não acentuados **/
            string[] acentos = new string[] { "ç", "Ç", "á", "é", "í", "ó", "ú", "ý", "Á", "É", "Í", "Ó", "Ú", "Ý", "à", "è", "ì", "ò", "ù", "À", "È", "Ì", "Ò", "Ù", "ã", "õ", "ñ", "ä", "ë", "ï", "ö", "ü", "ÿ", "Ä", "Ë", "Ï", "Ö", "Ü", "Ã", "Õ", "Ñ", "â", "ê", "î", "ô", "û", "Â", "Ê", "Î", "Ô", "Û" };
            string[] semAcento = new string[] { "c", "C", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "Y", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U", "a", "o", "n", "a", "e", "i", "o", "u", "y", "A", "E", "I", "O", "U", "A", "O", "N", "a", "e", "i", "o", "u", "A", "E", "I", "O", "U" };

            for (int i = 0; i < acentos.Length; i++)
            {
                str = str.Replace(acentos[i], semAcento[i]);
            }
            /** Troca os caracteres especiais da string por "" **/
            string[] caracteresEspeciais = { "¹", "²", "³", "£", "¢", "¬", "º", "¨", "\"", "'", ".", ",", "-", ":", "(", ")", "ª", "|", "\\\\", "°", "_", "@", "#", "!", "$", "%", "&", "*", ";", "/", "<", ">", "?", "[", "]", "{", "}", "=", "+", "§", "´", "`", "^", "~" };

            for (int i = 0; i < caracteresEspeciais.Length; i++)
            {
                str = str.Replace(caracteresEspeciais[i], "");
            }

            /** Troca os caracteres especiais da string por " " **/
            str = Regex.Replace(str, @"[^\w\.@-]", " ",
                                RegexOptions.None, TimeSpan.FromSeconds(1.5));

            return str.Trim();
        }
        private void button21_Click(object sender, EventArgs e)
        {
            string x = ChangeString("raimundo nonato mendonça dutra");

            StreamReader reader = new StreamReader(@"c:\temp\Bosch_26");
            StreamWriter writer = new StreamWriter(@"c:\temp\CompWFMSAPBIS.csv");
            try
            {
                string line = null;
                BSPersonsInfoFEMSA personsFemsa = new BSPersonsInfoFEMSA();
                List<BSPersonsInfoFEMSA> listFemsa = new List<BSPersonsInfoFEMSA>();
                int currentline = 1;
                string persno = null;
                string nomeWFM = null;
                string nomeSAP = null;
                string reSAP = null;
                string cardnoSAP = null;
                string reBIS = null;
                string nomeBIS = null;
                string cardnoBIS = null;
                string statusSAP = null;

                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckWFMBISSAP.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Carregando a lista do SAP.");
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split('|');

                    persno = arr[0].TrimStart(new Char[] { '0' });
                    string nome = arr[1].Replace("'", "");
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
                        NOME = nome.Replace("'", "").ToLower(),
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
                    listFemsa.Add(personsFemsa);
                }
                reader.Close();
                reader = null;

                this.LogTask.Information(String.Format("Lista do SAP carregada com sucesso. Total: {0} registros", listFemsa.Count.ToString()));

                writer.WriteLine("PERSID;PERSNO;NOME WFM;NOME SAP;CARDNO SAP;NOME BIS;CARDNO BIS;COMP. WFMxBIS;COMP. WFMxSAP;COMP. CARTAO;STATUS");

                this.LogTask.Information("Pesquisando WFM.");
                string sql = String.Format("select distinct data, be.re, be.nome, d30.cod_situacao, status, descricao_situacao, site, horario_ent, horario_sai from [10.153.68.133].sisqualwfm.[dbo].[BOSCH_EmpregadosSituacao_D+30] d30 " +
                "inner join[10.153.68.133].sisqualwfm.dbo.BOSCH_Empregados be on be.re = d30.re inner join[10.153.68.133].sisqualwfm.dbo.BOSCH_TIPOS_SITUACAO bs on d30.COD_SITUACAO = bs.cod_situacao " +
                "where convert(date, data, 103) = '{0}' and site not in ('JURUBATUBA', 'JURUBATUBA DIRETORIA', 'JURUBATUBA OPERAÇÕES')", DateTime.Now.ToString("MM/dd/yyyy"));

                using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                {
                    if (table.Rows.Count > 0)
                    {
                        this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                        foreach (DataRow row in table.Rows)
                        {
                            this.LogTask.Information("{0} de {1}. Pesquisando RE: {2}, Nome: {3}", (currentline++).ToString(), table.Rows.Count.ToString(), row["re"].ToString(), row["nome"].ToString());
                            persno = row["re"].ToString();
                            nomeWFM = row["nome"].ToString().ToLower();

                            this.LogTask.Information("Pesquisando no SAP.");
                            List<BSPersonsInfoFEMSA> lr = listFemsa.FindAll(delegate (BSPersonsInfoFEMSA sap) { return sap.PERSNO.Equals(persno); });
                            if (lr != null && lr.Count > 0)
                            {
                                this.LogTask.Information("{0} registros encontrados no SAP.", lr.Count.ToString());
                                reSAP = lr[0].PERSNO;
                                nomeSAP = !String.IsNullOrEmpty(lr[0].NOME) ? lr[0].NOME.ToLower() : "SEM NOME";
                                cardnoSAP = !String.IsNullOrEmpty(lr[0].CARDNO) ? lr[0].CARDNO : "SEM CARTAO";
                                statusSAP = lr[0].STATUSPROPRIO;
                            }
                            else
                            {
                                nomeSAP = "";
                                cardnoSAP = "";
                                statusSAP = "";
                                this.LogTask.Information("Nenhum registro encontrado no SAP.");
                            }

                            this.LogTask.Information("Pesquisando no BIS.");
                            using (DataTable tableBIS = this.bisACEConnection.loadDataTable(String.Format("select per.persid, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), persno, cardno from bsuser.persons per " +
                                " left outer join bsuser.cards cd on cd.persid = per.persid " +
                                "where per.status = 1 and cd.status != 0 and persno = '{0}' and PERSCLASSID = 'FF0000E500000001'", persno)))
                            {
                                if (tableBIS != null && tableBIS.Rows.Count > 0)
                                {
                                    this.LogTask.Information("{0} registros encontrados no BIS.", tableBIS.Rows.Count.ToString());
                                    foreach (DataRow rowBIS in tableBIS.Rows)
                                    {
                                        reBIS = rowBIS["persno"].ToString();
                                        nomeBIS = !String.IsNullOrEmpty(rowBIS["nome"].ToString()) ? rowBIS["nome"].ToString().ToLower() : "SEM NOME";
                                        cardnoBIS = !String.IsNullOrEmpty(rowBIS["cardno"].ToString()) ? rowBIS["cardno"].ToString().TrimStart(new Char[] { '0' }) : "SEM CARTAO";
                                        writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}", rowBIS["persid"].ToString(), persno, ChangeString(nomeWFM), ChangeString(nomeSAP), cardnoSAP, ChangeString(nomeBIS), cardnoBIS, ChangeString(nomeBIS).ToLower().Equals(ChangeString(nomeWFM).ToLower()), ChangeString(nomeSAP).ToLower().Equals(ChangeString(nomeWFM).ToLower()), cardnoBIS.Equals(cardnoSAP), statusSAP));
                                    }
                                }
                                else
                                {
                                    nomeBIS = "";
                                    cardnoBIS = "";
                                    this.LogTask.Information("Nenhum registro encontrado no BIS.");
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}", "NO PERSID", persno, ChangeString(nomeWFM), ChangeString(nomeSAP), cardnoSAP, ChangeString(nomeBIS), cardnoBIS, ChangeString(nomeBIS).ToLower().Equals((nomeWFM).ToLower()), ChangeString(nomeSAP).ToLower().Equals(nomeWFM.ToLower()), cardnoBIS.Equals(cardnoSAP), statusSAP));
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            string line = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\UnblockPeople.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Lendo o arquivo com os desbloqueios.");
                reader = new StreamReader(@"c:\temp\femsadesbloqueio.csv");

                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\femsadesbloqueio.csv");

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string sql = null;
                while ((line = reader.ReadLine()) != null)
                {
                    this.LogTask.Information(String.Format("{0} de {1}. PERSID: {2}", (currentline++).ToString(), totallines.ToString(), line));

                    this.LogTask.Information(String.Format("Carregando o colaborador: Persid: {0}", line));
                    BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                    if (persons.Load(line))
                    {
                        this.LogTask.Information("Desbloqueando a pessoa.");
                        if (persons.RemoveAllBLocks(this.bisACEConnection, line) == BS_ADD.SUCCESS)
                            this.LogTask.Information("Pessoa desbloqueada com sucesso.");
                        else
                            this.LogTask.Information("Erro desbloqueando a pessoa.");

                        persons.AUTHFROM = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0);
                        persons.AUTHUNTIL = DateTime.MinValue;
                        persons.Update();
                        persons.Dispose();
                    }

                    //sql = String.Format("select persid from bsuser.persons where status = 1 and persclass = 'E' and persno = '{0}'", line);
                    //BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                    //using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    //{
                    //    if (table.Rows.Count > 0)
                    //    {
                    //        this.LogTask.Information(String.Format("Carregando o colaborador: Persid: {0}", table.Rows[0]["persid"].ToString()));
                    //        if (persons.Load(table.Rows[0]["persid"].ToString()))
                    //        {
                    //            this.LogTask.Information("Desbloqueando a pessoa.");
                    //            if (persons.RemoveAllBLocks(this.bisACEConnection, table.Rows[0]["persid"].ToString()) == BS_ADD.SUCCESS)
                    //                this.LogTask.Information("Pessoa desbloqueada com sucesso.");
                    //            else
                    //                this.LogTask.Information("Erro desbloqueando a pessoa.");

                    //            persons.AUTHFROM = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0); 
                    //            persons.AUTHUNTIL = DateTime.MinValue;
                    //            persons.Update();
                    //        }
                    //    }
                    //}
                }

                this.Cursor = Cursors.WaitCursor;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {

        }

        private void button24_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            StreamWriter writer = null;
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckRETrueCardFalse.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                reader = new StreamReader(@"c:\temp\checkretruecardfalse.csv");
                writer = new StreamWriter(@"c:\temp\changecardSAP_2401.csv");

                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\checkretruecardfalse.csv");

                this.ReadParameters();
                this.bisLOGConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string line = null;
                string sql = null;
                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8}", "persid", "Nome BIS", "RE BIS", "Cardno BIS", "Nome SAP", "RE SAP", "Cardno SAP", "Cardno Ultimo", "Data Ultimo"));
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    string persid = arr[0];
                    string nomeBIS = arr[1];
                    string reBIS = arr[2];
                    string cardBIS = arr[3];
                    string nomeSAP = arr[4];
                    string reSAP = arr[5];
                    string cardSAP = arr[6];
                    this.LogTask.Information(String.Format("{0} de {1}. Verificando Nome: {2}, RE: {3}, PERSID: {4}, CardBIS: {5}, CardSAP: {6}", (currentline++).ToString(), totallines.ToString(), nomeBIS, reBIS, persid, cardBIS, cardSAP));

                    using (DataTable table = this.bisLOGConnection.loadDataTable(String.Format("set dateformat 'dmy' exec spLastUsedCardID_PERSNO '{0}'", reBIS)))
                    {
                        if (table.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("Ultimo cartao encontrado. Cardid: {0}, Data: {1}", table.Rows[0]["cardid"].ToString(), DateTime.Parse(table.Rows[0]["dataacesso"].ToString()).ToString("dd/MM/yyyy HH:mm:ss")));
                            this.LogTask.Information("Recuperando informacoes do cartao.");
                            using (DataTable tableCard = this.bisACEConnection.loadDataTable(String.Format("select cardid, cardno, sitecode = convert(int,convert(varbinary(4), CODEDATA)) from bsuser.cards where cardid = '{0}'", table.Rows[0][0].ToString())))
                            {
                                if (tableCard != null && tableCard.Rows.Count > 0)
                                {
                                    this.LogTask.Information(String.Format("Informacoes encontradas. Sitecode: {0}, Cardno: {1}", tableCard.Rows[0]["sitecode"].ToString(), tableCard.Rows[0]["cardno"].ToString()));

                                    this.LogTask.Information("Verificando todos os cartoes do RE.");
                                    using (DataTable tableRECard = this.bisACEConnection.loadDataTable(String.Format("select per.persid, cardid, cardno from bsuser.persons per inner join bsuser.cards cd on cd.persid = per.persid where persno = '{0}' " +
                                        "and per.status = 1 and cd.status = 1", reBIS)))
                                    {
                                        if (tableRECard != null && tableRECard.Rows.Count > 0)
                                        {
                                            this.LogTask.Information(String.Format("{0} cartoes encontrados.", tableRECard.Rows.Count.ToString()));
                                            //BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                                            foreach (DataRow row in tableRECard.Rows)
                                            {
                                                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8}", row["persid"].ToString(), nomeBIS, reBIS, row["cardno"].ToString(), nomeSAP, reSAP, cardSAP, tableCard.Rows[0]["cardno"].ToString(), DateTime.Parse(table.Rows[0]["dataacesso"].ToString()).ToString("dd/MM/yyyy HH:mm:ss")));

                                                //if (row["cardno"].ToString().Equals(tableCard.Rows[0]["cardno"].ToString()))
                                                //{
                                                //    this.LogTask.Information(String.Format("Cartoes iguais {0} = {1}. Proximo cartao.", row["cardno"].ToString(), tableCard.Rows[0]["cardno"].ToString()));
                                                //    continue;
                                                //}
                                                //else
                                                //    this.LogTask.Information(String.Format("Cartoes diferentes {0} != {1}. Excluir cartao.", row["cardno"].ToString(), tableCard.Rows[0]["cardno"].ToString()));

                                                //this.LogTask.Information("Carregando a pessoa.");
                                                //if (per.Load(row["persid"].ToString()))
                                                //{
                                                //    this.LogTask.Information("Pessoa carregada com sucesso.");
                                                //    this.LogTask.Information(String.Format("Excluindo o cartao. Cardid: {0}, Cardno: {1}", row["cardno"].ToString(), row["cardid"].ToString()));
                                                //    if (per.DetachCard(row["cardid"].ToString()) == BS_ADD.SUCCESS)
                                                //    {
                                                //        this.LogTask.Information("Cartao removido com sucesso.");
                                                //    }
                                                //    else
                                                //    {
                                                //        this.LogTask.Information("Erro ao remover o cartao.");
                                                //    }
                                                //}
                                            }
                                        }
                                    }
                                    //this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5}", nome, persno, persid, table.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                    //writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", nome, persno, persid, table.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                }
                                else
                                {
                                    this.LogTask.Information("Cartao nao encontrado.");
                                }
                            }
                        }
                        else
                        {
                            using (DataTable tableCur = this.bisACEConnection.loadDataTable(String.Format("select accesstime from bsuser.CURRENTACCESSSTATE where persid = '{0}'", persid)))
                            {
                                if (tableCur != null && tableCur.Rows.Count > 0)
                                {
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8}", persid, nomeBIS, reBIS, cardBIS, nomeSAP, reSAP, cardSAP, "SEM CARTAO", tableCur.Rows[0]["accesstime"].ToString()));
                                }
                                else
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8}", persid, nomeBIS, reBIS, cardBIS, nomeSAP, reSAP, cardSAP, "SEM CARTAO", "SEM DATA"));
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;

                reader.Close();
                writer.Close();

                this.BISManager.Disconnect();
                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\checkname_wfmbis.csv");
            StreamWriter writer = new StreamWriter(@"c:\temp\checkname_result.csv");
            try
            {
                string line = null;

                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckName_WFMBIS.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string persid = null;
                string persno = null;
                string nameWFM = null;
                string nameBIS = null;

                writer.WriteLine("PERSID;PERSNO;NOME WFM;NOME BIS;");

                while ((line = reader.ReadLine()) != null)
                {
                    this.LogTask.Information(line);

                    string[] arr = line.Split(';');

                    persid = arr[0];
                    persno = arr[1];
                    nameWFM = arr[2];
                    nameBIS = arr[5];

                    string fWFM = null;
                    string fBIS = null;
                    string lWFM = null;
                    string lBIS = null;

                    BSCommon.ParseName(nameWFM, out fWFM, out lWFM);
                    BSCommon.ParseName(nameBIS, out fBIS, out lBIS);

                    if (fWFM.ToLower().Equals(fBIS.ToLower()))
                        writer.WriteLine(String.Format("{0};{1};{2};{3}", persid, persno, nameWFM, nameBIS));
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                reader.Close();
                writer.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\Bosch_26");
            try
            {
                string line = null;
                BSPersonsInfoFEMSA personsFemsa = new BSPersonsInfoFEMSA();
                List<BSPersonsInfoFEMSA> listFemsa = new List<BSPersonsInfoFEMSA>();
                int currentline = 1;
                string persid = null;
                string persno = null;
                string nomeWFM = null;
                string nomeSAP = null;
                string reSAP = null;
                string cardnoSAP = null;
                string reBIS = null;
                string nomeBIS = null;
                string cardnoBIS = null;
                string statusSAP = null;

                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckWFMBISSAP.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Carregando a lista do SAP.");
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split('|');

                    persno = arr[0].TrimStart(new Char[] { '0' });
                    string nome = arr[1].Replace("'", "");
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
                        NOME = nome.Replace("'", "").ToLower(),
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
                    listFemsa.Add(personsFemsa);
                }
                reader.Close();
                reader = null;

                this.LogTask.Information(String.Format("Lista do SAP carregada com sucesso. Total: {0} registros", listFemsa.Count.ToString()));

                reader = new StreamReader(@"c:\temp\checkname_result.csv");
                reader.ReadLine();
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');

                    persid = arr[0];
                    persno = arr[1];

                    List<BSPersonsInfoFEMSA> lr = listFemsa.FindAll(delegate (BSPersonsInfoFEMSA sap) { return sap.PERSNO.Equals(persno); });

                    nomeBIS = lr[0].NOME;
                    string firstname = null;
                    string lastname = null;
                    BSCommon.ParseName(nomeBIS, out firstname, out lastname);

                    this.LogTask.Information("Atualizando o nome.");
                    string sql = String.Format("update bsuser.persons set firstname = '{0}', lastname = '{1}' where persid = '{2}'", firstname, lastname, persid);
                    this.bisACEConnection.executeProcedure(sql);
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                reader.Close();

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            StreamReader reader = new StreamReader(@"c:\temp\bisfromsap.csv");
            this.Cursor = Cursors.WaitCursor;
            this.LogTask = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckWFMBISSAP.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();

            this.ReadParameters();
            this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
            string line = null;

            while ((line = reader.ReadLine()) != null)
            {
                string[] arr = line.Split(';');

                if (!this.bisACEConnection.executeProcedure(String.Format("update bsuser.persons set persno = '{0}' where persid = '{1}'", arr[1], arr[0])))
                    MessageBox.Show("Erro.");
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\DeletNOCARD.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                string sql = "select cardid, persno, per.persid, cd.persid, cardno, cur.ACCESSTIME from bsuser.persons per left outer join bsuser.cards cd on cd.persid = per.persid " +
                    "left outer join bsuser.CURRENTACCESSSTATE cur on cur.persid = per.persid where per.status = 1 and cd.persid is null and persclass = 'E' and PERSCLASSID = 'FF0000E500000001' and ACCESSTIME is null";

                using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                {
                    BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                    foreach(DataRow row in table.Rows)
                    {
                        this.LogTask.Information("Excluindo o regsitro. PERSID: {0}, PERSNO: {1}", row["persid"].ToString(), row["persno"].ToString());
                        if (!String.IsNullOrEmpty(row["cardid"].ToString()))
                        {
                            this.LogTask.Information("Cartao existente.");
                            continue;
                        }
                        if (per.DeletePerson(row["persid"].ToString()) == BS_ADD.SUCCESS)
                            this.LogTask.Information("Registro excluído com sucesso.");
                        else
                            this.LogTask.Information("Erro ao excluir o registro.");
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information("Descontecando o BIS.");
                if (this.BISManager.Disconnect())
                    this.LogTask.Information("BIS desconectado com sucesso.");
                else
                    this.LogTask.Information("Erro ao desconectar com o BIS.");

                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            this.LogTask = new LoggerConfiguration()
    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3Check.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
    .CreateLogger();

            this.ReadParameters();
            this.BISACESQL.SQLHost = "172.16.0.86\\BIS_ACE";
            this.BISACESQL.SQLUser = "sa";
            this.BISACESQL.SQLPwd = "S$iG3L9a@n";
            this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
            DataTable table = BSSQLPersons.GetPersonsLight(this.bisACEConnection, "persid", "001375E0E5242295");
            var retval = Global.ConvertDataTable<BSPersonsInfo>(table);

        }
    }
}
