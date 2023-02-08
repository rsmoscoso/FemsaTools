using AMSLib;
using AMSLib.ACE;
using AMSLib.Manager;
using AMSLib.SQL;
using HzBISCommands;
using HzLibConnection.Data;
using Newtonsoft.Json;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FemsaTools
{
    public partial class Form2 : Form
    {

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
        public HzConexao bisConnection { get; set; }
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

                //this.LogTask.Information("Lendo Arquivo de compartilhamento...");
                //reader = new StreamReader(@"c:\horizon\config\credential.json");
                //sql = reader.ReadToEnd();
                //reader.Close();

                //if (String.IsNullOrEmpty(sql))
                //    throw new Exception("Nao ha arquivo com os dados do compartilhamento.");

                //this.femsaCredential = JsonConvert.DeserializeObject<FemsaNetworkCredential>(sql);
                //this.femsaCredential.SharingFolder = @"\\kofbrsmcappd00.sa.kof.ccf\Interfaces\RRHH";
                //this.LogTask.Information(String.Format("Credential --> User: {0}, Pwd: {1}, Domain: {2}, SharingFolder: {3}, FileName: {4}", this.femsaCredential.UserName,
                //    this.femsaCredential.Password, this.femsaCredential.Domain, this.femsaCredential.SharingFolder, this.femsaCredential.FileName));

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
        public Form2()
        {
            InitializeComponent();

            this.LogTask = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\BlockTaskLog.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();

            this.LogTaskSolo = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\BlockTaskLogSolo.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();

            this.LogResult = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\BlockResultLog.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();

            this.LogResultSolo = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\BlockResultLogSolo.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.ReadParameters();

            CultureInfo ptBR = new CultureInfo("pt-BR", true);
            HzConexao conACE = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
            HzConexao conBIS = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "acedb", "System.Data.SqlClient");

            FemsaBlocks femsaBlocks = new FemsaBlocks(this.LogTaskSolo, this.LogResultSolo, this.BISACESQL, this.BISSQL, this.BISManager, ptBR);
            femsaBlocks.CheckBlocks(null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.ReadParameters();

            CultureInfo ptBR = new CultureInfo("pt-BR", true);
            HzConexao conACE = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
            HzConexao conBIS = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "acedb", "System.Data.SqlClient");

            FemsaBlocks femsaBlocks = new FemsaBlocks(this.LogTaskSolo, this.LogResultSolo, this.BISACESQL, this.BISSQL, this.BISManager, ptBR);
            femsaBlocks.CheckExceptions(this.BISManager);

        }

        private void button2_Click(object sender, EventArgs e)
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

        private void btnExecute_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtRE.Text))
            {
                MessageBox.Show("Digite um RE!", "Horizon", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            this.LogTask = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\ImportTaskLogRE.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();
            this.LogResult = new LoggerConfiguration()
                .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Block\\ImportResultLogRE.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                .CreateLogger();
            this.ReadParameters();

            CultureInfo ptBR = new CultureInfo("pt-BR", true);
            HzConexao conACE = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");
            HzConexao conBIS = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "acedb", "System.Data.SqlClient");

            FemsaBlocks femsaBlocks = new FemsaBlocks(this.LogTask, this.LogResult, this.BISACESQL, this.BISSQL, this.BISManager, ptBR);
            femsaBlocks.CheckBlocks(txtRE.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\ImportProprios.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();

                StreamReader reader = new StreamReader(@"c:\temp\Bosch_8");
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
                femsa.ImportList(list);
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information("Rotina finalizada..");
            }

        }
    }
}
