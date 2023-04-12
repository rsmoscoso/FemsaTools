using AMSLib;
using AMSLib.ACE;
using AMSLib.Common;
using AMSLib.Manager;
using AMSLib.SQL;
using FemsaTools.AMS;
using FemsaTools.SG3;
using HzBISCommands;
using HzLibConnection.Data;
using HzLibWindows.Common;
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
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Runtime.ConstrainedExecution;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static FemsaTools.FEMSAImport;
using static HzBISCommands.BSPersonsInfo;
using static System.Net.Mime.MediaTypeNames;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

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
        /// Log das tarefas da importacao do SG3.
        /// </summary>
        public ILogger LogSG3Task { get; set; }
        /// <summary>
        /// Resultado dos bloqueios/desbloqueios.
        /// </summary>
        public ILogger LogResult { get; set; }
        /// <summary>
        /// Resultado dos bloqueios/desbloqueios.
        /// </summary>
        public ILogger LogResultSolo { get; set; }
        /// <summary>
        /// Resultado da da importacao do SG3..
        /// </summary>
        public ILogger LogSG3Result { get; set; }
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
        /// <summary>
        /// Conexão com o banco de dados de Eventos do BIS.
        /// </summary>
        public HzConexao bisConnection { get; set; }
        /// <summary>
        /// Dados para o compartilhamento do arquivo de importação.
        /// </summary>
        private FemsaNetworkCredential femsaCredential { get; set; }
        /// <summary>
        /// Lista dos hosts do SG3.
        /// </summary>
        private List<SG3Host> SG3Hosts { get; set; }
        /// <summary>
        /// Aborta a execução de uma tarefa.
        /// </summary>
        private bool abort { get; set; }

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

                //this.SG3Hosts = new List<SG3Host>();
                //this.SG3Hosts.Add(new SG3Host() { Host = "http://api2.executiva.adm.br", Command = "femsabrasil/v1/alocacao/liberada", Username = "webservice.femsabrasil", Password = "T8yP@2jK$4r" });
                //this.SG3Hosts.Add(new SG3Host() { Host = "http://api2.executiva.adm.br", Command = "femsat1/v1/alocacao/liberada", Username = "webservice.femsat1", Password = "T8yP@2jK$4r" });
                //this.SG3Hosts.Add(new SG3Host() { Host = "http://api2.executiva.adm.br", Command = "femsa/v1/alocacao/liberada", Username = "webservice.femsa", Password = "T8yP@2jK$4r" });

                //StreamWriter writer = new StreamWriter(@"c:\horizon\config\sg3.json");
                //writer.WriteLine(JsonConvert.SerializeObject(this.SG3Hosts));
                //writer.Close();

                this.LogTask.Information("Lendo SG3...");
                reader = new StreamReader(@"c:\horizon\config\sg3.json");
                sql = reader.ReadToEnd();
                reader.Close();

                if (String.IsNullOrEmpty(sql))
                    throw new Exception("Nao ha arquivo de configuracao do SG3.");

                this.SG3Hosts = JsonConvert.DeserializeObject<List<SG3Host>>(sql);
                foreach(SG3Host host in this.SG3Hosts)
                {
                    this.LogTask.Information(String.Format("SG3 --> Host: {0}, Username: {1}, Password: {2}", host.Host,
                        host.Username, host.Password));
                }

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
                throw new Exception(ex.Message);
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

                this.Cursor = Cursors.Default;
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

                StreamReader reader = new StreamReader(@"c:\temp\Bosch");
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

        /// <summary>
        /// Retorna o tipo da pesquisa.
        /// </summary>
        /// <returns></returns>
        private string getInfoType()
        {
            if (rdbAccess.Checked)
                return "access";
            else if (rdbAuthorized.Checked)
                return "authoriza";
            else if (rdbNotYet.Checked)
                return "not yet";
            else
                return null;
        }

        /// <summary>
        /// Gera o csv do arquivo info.
        /// </summary>
        /// <param name="filename">Nome do arquivo.</param>
        private bool formatCSV(string filename)
        {
            StreamReader reader = null;
            try
            {
                this.Cursor = Cursors.WaitCursor;
                float length = 0;
                reader = new StreamReader(filename);
                while (reader.ReadLine() != null)
                    length++;
                reader.Close();
                this.abort = false;
                reader = new StreamReader(filename);
                string line = null;
                int pos = -1;
                string cardid = null;
                string deviceid = null;
                BSPersonsInfo person = new BSPersonsInfo();
                string[] vw = new string[8];
                ListViewItem item = null;
                string tipo = getInfoType();
                grpBar.Visible = true;
                grpBar.Text = String.Format("Arquivo: {0}", filename);
                pgBar.Refresh();
                float lineCount = 0;

                this.LogTask.Information(String.Format("Arquivo: {0}", filename));
                while ((line = reader.ReadLine()) != null)
                {
                    this.LogTask.Information(String.Format("{0} de {1}: {2}", (++lineCount).ToString(), length.ToString(), line));

                    pgBar.Value = (int)Math.Truncate((lineCount / length) * 100);
                    pgBar.Refresh();
                    if (lineCount % 1000 == 0)
                        lstData.Refresh();

                    System.Windows.Forms.Application.DoEvents();

                    if (this.abort)
                        break;

                    if (lineCount % 1000 == 0)
                        this.Refresh();

                    string[] arr = line.Trim().Split(' ');

                    if ((pos = arr[0].IndexOf("-")) <= 0 || arr.Length < 10)
                        continue;

                    try
                    {
                        string id = arr[0].Substring(0, pos);
                        DateTime date = DateTime.Parse(arr[1] + " " + arr[2]);
                        deviceid = arr[4];
                        cardid = arr[6];

                        // exibe apenas os eventos relativos ao acesso
                        if ((String.IsNullOrEmpty(tipo) && !String.IsNullOrEmpty(arr[6])) || (!String.IsNullOrEmpty(arr[7])
                            && !String.IsNullOrEmpty(arr[4]) && !String.IsNullOrEmpty(arr[6]) && line.IndexOf(tipo) > 0))
                        {
                            person = BSSQLPersons.GetPersonFromCardID(this.bisACEConnection, cardid);
                            if (person == null)
                                continue;
                            if (String.IsNullOrEmpty(txtREInfo.Text) || !String.IsNullOrEmpty(txtREInfo.Text) && person.DOCUMENT.Equals(txtREInfo.Text))
                            {
                                person.RESERVE1 = BSSQLDevices.GetDescription(this.bisACEConnection, deviceid);
                                if (person != null)
                                {
                                    vw[0] = arr[0].Substring(pos + 1, (arr[0].Length - (pos + 1))) + " " + arr[1];
                                    vw[1] = person.PERSID;
                                    vw[2] = person.PERSCLASS;
                                    vw[3] = person.DOCUMENT;
                                    vw[4] = person.NAME;
                                    vw[5] = person.CARDNO;
                                    vw[6] = person.RESERVE1;
                                    vw[7] = !String.IsNullOrEmpty(arr[7]) ? arr[7] : "Indefinido";
                                    item = new ListViewItem(vw);
                                    lstData.Items.Add(item);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.Cursor = Cursors.Default;
                        MessageBox.Show(ex.Message);
                    }
                }
                reader.Close();
                this.Cursor = Cursors.Default;

                return !this.abort;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message);
                return !this.abort;
            }
            finally
            {
                this.Cursor = Cursors.Default;
                if (reader != null)
                    reader.Close();
                pgBar.Visible = false;
                if (lstData.Items.Count > 0)
                    btnExport.Visible = true;
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                lstData.View = View.Details;

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                if (!chkType.Checked)
                {
                    if (openFileDialog1.ShowDialog(this) != DialogResult.OK)
                        return;

                    lstData.Items.Clear();
                    this.formatCSV(openFileDialog1.FileName);
                }
                else
                {
                    if (folderBrowserDialog1.ShowDialog() != DialogResult.OK)
                        return;

                    lstData.Items.Clear();
                    string path = folderBrowserDialog1.SelectedPath;
                    foreach (string filename in Directory.GetFiles(path))
                    {
                        if (!this.formatCSV(filename))
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                grpBar.Visible = false;
                if (lstData.Items.Count > 0)
                    btnExport.Visible = true;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            StreamWriter writer = null;
            try
            {
                if (this.saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;
                this.Cursor = Cursors.WaitCursor;
                writer = new StreamWriter(this.saveFileDialog1.FileName);
                writer.WriteLine("Data;Persid;Tipo;Documento;Nome;Cartao;Local;Mensagem");
                foreach (ListViewItem item in lstData.Items)
                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7}", item.SubItems[0].Text, item.SubItems[1].Text, item.SubItems[2].Text, item.SubItems[3].Text, 
                        item.SubItems[4].Text, item.SubItems[5].Text, item.SubItems[6].Text, item.SubItems[7].Text));

                writer.Close();
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                pgBar.Visible = false;
                if (lstData.Items.Count > 0)
                    btnExport.Visible = true;
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                using (DataTable table = BSSQLClients.GetClients(this.bisACEConnection, "geral", "null"))
                    Objetos.LoadCombo(cmbDivisao, table, "NAME", "CLIENTID", "NAME", true);
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string authstr = null;
                string persid = null;
                List<string> authString = new List<string>();
                List<BSAuthorizationInfo> auths = null;
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\ChangeDivision.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.Cursor = Cursors.WaitCursor;
                this.LogTask.Information(String.Format("Pesquisando o PERSNO: {0}", txtREDivisao.Text));
                using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select persid from bsuser.persons where persno = '{0}' and status = 1", txtREDivisao.Text)))
                {
                    if (table == null || table.Rows.Count < 1)
                        throw new Exception("Colaborador não encontrado!");

                    persid = table.Rows[0]["persid"].ToString();
                    auths = BSSQLPersons.GetAllAuthorizations(this.bisACEConnection, table.Rows[0]["persid"].ToString());
                    if (auths == null || auths.Count < 1)
                        this.LogTask.Information("Nao ha autorizacoes para esse colaborador.");
                    else
                    {
                        authString.Clear();
                        this.LogTask.Information(String.Format("{0} Autorizacoes encontradas.", table.Rows.Count.ToString()));
                        foreach (BSAuthorizationInfo auth in auths)
                        {
                            authString.Add(auth.AUTHID);
                        }
                    }
                }

                int totalBefore = auths != null ? auths.Count : 0;
                int totalAfter = -1;

                using (BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection))
                {
                    this.LogTask.Information("Carregando a pessoa.");
                    if (persons.Load(txtREDivisao.Text, BS_STATUS.ACTIVE))
                    {
                        this.LogTask.Information("Pessoa carregada com sucesso.");
                        this.LogTask.Information("Alterando a divisao.");
                        if (persons.ChangeClient(((HzItem)cmbDivisao.SelectedItem).value) == BS_SAVE.SUCCESS)
                        {
                            this.LogTask.Information("Divisao alterada com sucesso.");
                            this.LogTask.Information("Atribuindo as autorizacoes");
                            if (persons.SetAuthorization(authString.ToArray()) == BS_ADD.SUCCESS)
                                this.LogTask.Information("Autorizacoes atribuidas com sucesso.");

                            auths = BSSQLPersons.GetAllAuthorizations(this.bisACEConnection, persid);
                            totalAfter = auths != null ? auths.Count : 0;

                            if (totalAfter != totalBefore)
                                throw new Exception(String.Format("Total Antes: {0}, Total Depois: {1}. Erro na alteracao. Rotina interrompida.", totalBefore.ToString(), totalAfter.ToString()));

                            this.bisACEConnection.executeProcedure(String.Format("update bsuser.acpersons set clientid = '{0}' where persid = '{1}'", ((HzItem)cmbDivisao.SelectedItem).value,
                                persons.GetPersonId()));

                            this.LogTask.Information(String.Format("Total Antes: {0}, Total Depois: {1}. Alteracao concluida com sucesso.", totalBefore.ToString(), totalAfter.ToString()));
                        }
                        else
                            this.LogTask.Information("Erro ao alterar a divisao.");
                    }
                    else
                        this.LogTask.Information("Erro ao carregar a pessoa.");

                    MessageBox.Show("Alteração concluída com sucesso!");
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

        private bool IsOlderThan30(Liberacao l)
        {
            return l.liberado.Equals("S") && l.status_alocacao.Equals("S");
        }
        private async void button7_Click(object sender, EventArgs e)
        {
            //StreamWriter writer = new StreamWriter(@"c:\temp\sg3tipo.csv");
            BSPersons persons = null;
            try
            {
                this.LogSG3Task = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\ImportSG3TaskLogSolo.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.LogSG3Result = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\ImportSG3ResultLogSolo.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                persons = new BSPersons(this.BISManager, this.bisACEConnection);

                SG3Import import = new SG3Import(this.LogSG3Task, this.LogSG3Result, this.BISACESQL, this.BISSQL, this.BISManager, this.SG3Hosts);
                import.Import();
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogSG3Task.Information(String.Format("Erro: {0}", ex.Message));
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

        private void button8_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            try
            {
                string line = null;
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AddCompany.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Lendo o arquivo com as empresas.");
                if (openFileDialog1.ShowDialog() == DialogResult.Cancel) throw new Exception("Nenhum arquivo lido.");

                this.Cursor = Cursors.WaitCursor;

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");


                reader = new StreamReader(openFileDialog1.FileName);
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');
                    this.LogTask.Information(String.Format("CNPJ: {0}, Nome: {1}", arr[0], arr[1]));
                    string id = BSSQLCompanies.GetID(this.bisACEConnection, arr[1]);
                    if (String.IsNullOrEmpty(id))
                    {
                        this.LogTask.Information("Nova Empresa");
                        ACECompanies cmp = new ACECompanies(this.BISManager.m_aceEngine);
                        cmp.NAME = arr[1].ToUpper();
                        cmp.COMPANYNO = arr[1].ToUpper();
                        cmp.REMARKS = arr[0];
                        if (cmp.Add() == API_RETURN_CODES_CS.API_SUCCESS_CS)
                            this.LogTask.Information("Empresa gravada com sucesso.");
                        else
                            this.LogTask.Information("Erro ao gravar a empresa.");
                    }
                }

                this.Cursor = Cursors.Default;
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

        private void button9_Click(object sender, EventArgs e)
        {
            StreamWriter writer = new StreamWriter(@"c:\temp\Persclassid.csv");

            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckWFMBIS.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                using (DataTable table = this.bisACEConnection.loadDataTable("select distinct re, nome from [10.153.68.133].sisqualwfm.dbo.BOSCH_Empregados be"))
                {
                    foreach(DataRow row in table.Rows)
                    {
                        using (DataTable tblBIS = this.bisACEConnection.loadDataTable(String.Format("select persid, persclass, persclassid from bsuser.persons where persno = '{0}' and persclassid != 'FF0000E500000001'", row["re"].ToString())))
                        {
                            foreach (DataRow r in tblBIS.Rows)
                                this.bisACEConnection.executeProcedure(String.Format("update bsuser.persons set persclassid = 'FF0000E500000001' where persid = '{0}'", r["persid"].ToString()));
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
                writer.Close();
                this.Cursor = Cursors.Default;
                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            StreamWriter writer = new StreamWriter(@"c:\temp\CardID_SG3.csv");
            try
            {
                this.Cursor = Cursors.WaitCursor;

                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\CheckLastUsedCardSG3.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisConnection = new HzConexao(this.BISSQL.SQLHost, this.BISSQL.SQLUser, this.BISSQL.SQLPwd, "BISEventLog", "System.Data.SqlClient");
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                int currentline = 1;
                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", "persid", "re", "nome", "data", "cardid", "cardno", "sitecode"));
                using (DataTable table = this.bisACEConnection.loadDataTable(String.Format("select per.persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), accesstime from bsuser.persons per " +
                    "left outer join bsuser.CURRENTACCESSSTATE cur on cur.persid = per.persid where per.status = 1 and persclass = 'E' and persclassid != 'FF0000E500000001'")))
                {
                    this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                    foreach (DataRow row in table.Rows)
                    {
                        this.LogTask.Information("{0} de {1}. Pesquisando PERSID: {2}, RE: {3}, Nome: {4}", (currentline++).ToString(), table.Rows.Count.ToString(), row["persid"].ToString(), row["persno"].ToString(), row["nome"].ToString());
                        string persid = row["persid"].ToString();
                        string persno = row["persno"].ToString();
                        string nome = row["nome"].ToString();
                        string data = row["accesstime"].ToString();
                        if (String.IsNullOrEmpty(data))
                        {
                            this.LogTask.Information("Nao ha data de ultimo acesso.");
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", nome, persno, persid, "naoencontrado", "naoencontrado", "naoencontrado", "naoencontrado"));
                            continue;
                        }
                        else if (!String.IsNullOrEmpty(data) && DateTime.Parse(data).Subtract(DateTime.Parse("01/01/2023")).Days < 0)
                        {
                            this.LogTask.Information("Data Inferior a janeiro de 2023.");
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", nome, persno, persid, data, "inferior0123", "inferior0123", "inferior0123"));
                            continue;
                        }

                        this.LogTask.Information(String.Format("Verificando Nome: {0}, RE: {1}, PERSID: {2}", nome, persno, persid));

                        using (DataTable tableCD = this.bisConnection.loadDataTable(String.Format("set dateformat 'dmy' exec spLastUsedCardID_PERSID '{0}'", persid)))
                        {
                            if (tableCD.Rows.Count > 0)
                            {
                                using (DataTable tableCard = this.bisACEConnection.loadDataTable(String.Format("select cardid, cardno, sitecode = convert(int,convert(varbinary(4), CODEDATA)) from bsuser.cards where cardid = '{0}'", tableCD.Rows[0][0].ToString())))
                                {
                                    if (tableCard.Rows.Count > 0)
                                    {
                                        this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5};{6}", nome, persno, persid, data, tableCard.Rows[0]["cardid"].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                        writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", nome, persno, persid, data, table.Rows[0][0].ToString(), tableCard.Rows[0]["cardno"].ToString(), tableCard.Rows[0]["sitecode"].ToString()));
                                    }
                                    else
                                    {
                                        this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5};{6}", nome, persno, persid, data, "naoencontrado", "naoencontrado", "naoencontrado"));
                                        writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", nome, persno, persid, data, "naoencontrado", "naoencontrado", "naoencontrado"));
                                    }
                                }
                            }
                            else
                            {
                                this.LogTask.Information("PERSID nao encontrado.");
                                writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", nome, persno, persid, "naoencontrado", "naoencontrado", "naoencontrado", "naoencontrado"));
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

        private void button11_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            string line = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\DeleteSG3.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                if (openFileDialog1.ShowDialog() == DialogResult.Cancel) throw new Exception("Nenhum arquivo lido.");

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information(String.Format("Lendo o arquivo com os colaboradores. Filename: {0}", openFileDialog1.FileName));
                reader = new StreamReader(openFileDialog1.FileName);

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string sql = null;
                BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');
                    string persno = arr[0];
                    string nome = arr[1];
                    this.LogTask.Information(String.Format("Persno: {0}, Nome: {1}", persno, nome));

                    sql = String.Format("select persid, nome = isnull(firstname, '') + ' ' + isnull(lastname, '') from bsuser.persons where persno = '{0}' and persclassid != 'FF0000E500000001' and status = 1", persno);

                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("Excluindo o colaborador: Persid: {0}", table.Rows[0]["persid"].ToString()));
                            if (persons.DeletePerson(table.Rows[0]["persid"].ToString()) == BS_ADD.SUCCESS)
                                this.LogTask.Information("Pessoa excluída com sucesso.");
                            else
                                this.LogTask.Information("Erro ao excluir a pessoa.");
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

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                string line = null;
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\DelteNoAccess.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.Cursor = Cursors.WaitCursor;

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");
                this.LogTask.Information("Pesquisando os que nao tiveram acesso.");
                using (DataTable table = this.bisACEConnection.loadDataTable("select distinct per.persid, per.persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), accesstime, PERSCLASSID from bsuser.persons per " +
                    "left outer join bsuser.CURRENTACCESSSTATE cur on cur.persid = per.persid where persclass = 'E' and per.status = 1 and accesstime is null"))
                {
                    if (table != null && table.Rows.Count > 0) 
                    {
                        int i = 1;
                        BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                        this.LogTask.Information(String.Format("{0} Registros encontrados.", table.Rows.Count.ToString()));
                        foreach (DataRow row in table.Rows)
                        { 
                            this.LogTask.Information(String.Format("{0} de {1}. Persid: {2}, Persno: {3}, Nome: {4}", (i++).ToString(), table.Rows.Count.ToString(),
                                row["persid"].ToString(), row["persno"].ToString(), row["nome"].ToString()));
                            this.LogTask.Information("Excluindo a pessoa.");
                            if (per.DeletePerson(row["persid"].ToString()) == BS_ADD.SUCCESS)
                                this.LogTask.Information("Exclusao realizada com sucesso.");
                            else
                                this.LogTask.Information("Erro ao excluir a pessoa.");
                        }
                    }
                }

                this.Cursor = Cursors.Default;
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

        private void button13_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            StreamWriter writer = null;
            List<FemsaInfo> femsaInfo = new List<FemsaInfo>();
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AjusteCPF.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                reader = new StreamReader(@"c:\temp\Bosch_2");
                writer = new StreamWriter(@"c:\temp\CheckCPF.csv");
                string line = null;
                BSPersonsInfoFEMSA personsFemsa = null;
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\Bosch_2");

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
                    this.LogTask.Information(String.Format("RE: {0}, Nome: {1}, CPF: {2}, Status: {3}", personsFemsa.PERSNO, personsFemsa.NOME, personsFemsa.CPF, personsFemsa.STATUSPROPRIO));

                    this.LogTask.Information("Pesquisando a existencia do colaborador.");
                    string sql = String.Format("select distinct per.persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), cli.Name from bsuser.ADDITIONALFIELDS fd " +
                        "inner join bsuser.persons per on per.persid = fd.persid inner join bsuser.clients cli on cli.clientid = per.clientid where FIELDDESCID = '00137EFFB5D874FE' and persclass = 'E' and status = 1 and fd.value = '{0}'", personsFemsa.CPF);
                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table != null && table.Rows.Count > 0)
                        {
                            foreach (DataRow row in table.Rows)
                            {
                                //writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6}", row["persid"].ToString(), row["nome"].ToString(), personsFemsa.NOME, row["persno"].ToString(), personsFemsa.PERSNO, personsFemsa.STATUSPROPRIO, row["name"].ToString()));
                                this.LogTask.Information(String.Format("{0} de {1}. Excluindo: Persid: {1}, Nome: {2}", row["persid"].ToString(), row["nome"].ToString()));
                                using (BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection))
                                {
                                    if (persons.DeletePerson(row["persid"].ToString()) == BS_ADD.SUCCESS)
                                        this.LogTask.Information("Pessoa excluida com sucesso.");
                                    else
                                        this.LogTask.Information("Erro ao excluir a pessoa.");
                                }
                            }
                        }
                    }
                }
                //reader.Close();
                //writer.Close();
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Rotina finalizada.");
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            string line = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\RemoveAuthorizaion.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
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

        private void button16_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            string line = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AjusteSAPBIS.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Lendo o arquivo com as autorizacoes.");
                if (openFileDialog1.ShowDialog() == DialogResult.Cancel) throw new Exception("Nao ha arquivo para leitura.");

                reader = new StreamReader(openFileDialog1.FileName);

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                string sql = null;
                string firsntame = null;
                string lastname = null;
                string clientid = null;
                BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');
                    string persid = arr[0];
                    string nome = arr[1];
                    string re = arr[2];
                    string divisao = arr[3];
                    string validacao = arr[4];

                    BSCommon.ParseName(nome, out firsntame, out lastname);

                    this.LogTask.Information(String.Format("Acao: {0}, Persid: {1}, RE: {2}, Nome: {3}, Divisao: {4}", validacao, persid, re, nome, divisao));
                    if (validacao.Equals("Colocar RE SAP"))
                    {
                        sql = String.Format("update bsuser.persons set persno = '{0}' where persid = '{1}'", re, persid);
                    }
                    else if (validacao.Equals("Corrigir Nome"))
                    {
                        sql = String.Format("update bsuser.persons set firstname = '{0}', lastname = '{1}' where persid = '{2}'", firsntame, lastname, persid);
                    }
                    else if (validacao.Equals("Corrigir Nome e RE SAP"))
                    {
                        sql = String.Format("update bsuser.persons set firstname = '{0}', lastname = '{1}', persno = '{2}' where persid = '{3}'", firsntame, lastname, re, persid);
                    }
                    else if (validacao.Equals("Corrigir Nome e RE SAP e divisão"))
                    {
                        this.LogTask.Information("Carregando a pessoa.");
                        if (per.Load(persid))
                        {
                            this.LogTask.Information("Pessoa carregada com sucesso.");
                            this.LogTask.Information(String.Format("Pesquisando divisao: {0}", divisao));
                            if (!String.IsNullOrEmpty(clientid = BSSQLClients.GetID(this.bisACEConnection, divisao)))
                            {
                                this.LogTask.Information(String.Format("Clientid: {0}", clientid));
                                if (per.ChangeClient(clientid) == BS_SAVE.SUCCESS)
                                    this.LogTask.Information("Divisao alterada com sucesso.");
                                else
                                    this.LogTask.Information("Erro ao alterar a divisao.");
                            }
                            else
                                this.LogTask.Information("Divisao nao encontrada.");
                        }
                        else
                            this.LogTask.Information("Erro ao carregar a pessoa.");

                        sql = String.Format("update bsuser.persons set firstname = '{0}', lastname = '{1}', persno = '{2}' where persid = '{3}'", firsntame, lastname, re, persid);
                    }
                    else if (validacao.Equals("Corrigir nome SAP e Divisão"))
                    {
                        this.LogTask.Information("Carregando a pessoa.");
                        if (per.Load(persid))
                        {
                            this.LogTask.Information("Pessoa carregada com sucesso.");
                            this.LogTask.Information(String.Format("Pesquisando divisao: {0}", divisao));
                            if (!String.IsNullOrEmpty(clientid = BSSQLClients.GetID(this.bisACEConnection, divisao)))
                            {
                                this.LogTask.Information(String.Format("Clientid: {0}", clientid));
                                if (per.ChangeClient(clientid) == BS_SAVE.SUCCESS)
                                    this.LogTask.Information("Divisao alterada com sucesso.");
                                else
                                    this.LogTask.Information("Erro ao alterar a divisao.");
                            }
                            else
                                this.LogTask.Information("Divisao nao encontrada.");
                        }
                        else
                            this.LogTask.Information("Erro ao carregar a pessoa.");

                        sql = String.Format("update bsuser.persons set firstname = '{0}', lastname = '{1}' where persid = '{2}'", firsntame, lastname, persid);
                    }

                    this.LogTask.Information("Atualizando o sql.");
                    if (this.bisACEConnection.executeProcedure(sql))
                        this.LogTask.Information("SQL atualizado com sucesso.");
                    else
                        this.LogTask.Information("Erro ao atualizar o sql");
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

        private void button14_Click_1(object sender, EventArgs e)
        {
            StreamWriter writer = null;
            StreamReader reader = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\AMS\\\\AMSTaskSolo.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                if (txtFileName.isBlank())
                    throw new Exception("Nome do arquivo em branco!");

                if (folderBrowserDialog1.ShowDialog() != DialogResult.OK)
                    return;

                this.LogTask.Information(String.Format("Iniciando a importacao do diretorio: {0}", folderBrowserDialog1.SelectedPath));

                string line = null;
                DateTime dateTime = DateTime.Now;
                string door = null;
                string cardno = null;
                string name = null;
                string persno = null;
                string company = null;
                int pos = 0;
                bool isDateTime = false;
                int count = 0;
                float length = 0;
                int lineCount = 0;
                bool noCompany = false;
                string lastLine = null;
                AMSImport import = new AMSImport(this.LogTask);

                foreach (string filename in Directory.GetFiles(folderBrowserDialog1.SelectedPath))
                {
                    this.LogTask.Information(String.Format("Arquivo: {0}", filename));

                    length = 0;
                    reader = new StreamReader(filename);
                    while (reader.ReadLine() != null)
                        length++;
                    reader.Close();

                    writer = new StreamWriter(txtFileName.Text, true);
                    reader = new StreamReader(filename);
                    while ((line = reader.ReadLine()) != null)
                    {
                        this.LogTask.Information(String.Format("{0} de {1}: {2}", (++lineCount).ToString(), length.ToString(), line));

                        writer = import.parseLine(reader, writer, line);
                    }
                    reader.Close();
                    writer.Close();
                }

                MessageBox.Show("Rotina finalizada.");
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (writer != null)
                    writer.Close();
                if (reader != null)
                    reader.Close();
                this.Cursor = Cursors.Default;

                this.LogTask.Information("Rotina finalizada.");
            }
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            this.abort = e.KeyCode == Keys.Escape;
                
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\ChangeToCommon.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                this.LogTask.Information("Lendo os terceiros fora da Common.");
                using (DataTable table = this.bisACEConnection.loadDataTable("select persid, persno, name = isnull(firstname, '') + ' ' + isnull(lastname, ''), Unidade = cli.Name from bsuser.persons per " +
                    " inner join bsuser.clients cli on cli.clientid = per.clientid where per.status = 1 and PERSCLASSID = '0013B4437A2455E1' and grade = 'SG3' and persclass = 'E' and cli.clientid != 'FF0000D300000002'"))
                {
                    this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                    int rowCount = 1;
                    BSPersons per = new BSPersons(this.BISManager, this.bisACEConnection);
                    foreach(DataRow row in table.Rows)
                    {
                        this.LogTask.Information(String.Format("{0} de {1}. Persid: {2}, Persno: {3}, Nome: {4}, Unidade: {5}",
                            (rowCount++).ToString(), table.Rows.Count.ToString(), row["persid"].ToString(), row["persno"].ToString(), row["name"].ToString(), row["Unidade"].ToString()));

                        if (row["persno"].ToString().Length < 11)
                        {
                            this.LogTask.Information("CPF invalido.");
                            continue;
                        }
                        this.LogTask.Information("Carregando a pessoa.");
                        if (per.Load(row["persid"].ToString()))
                        {
                            this.LogTask.Information("Pessoa carregada com sucesso.");
                            this.LogTask.Information("Alterando para Common.");
                            if (per.ChangeClient("FF0000D300000002") == BS_SAVE.SUCCESS)
                                this.LogTask.Information("Divisao alterada com sucesso.");
                            else
                                this.LogTask.Information("Erro ao alterar a divisao.");
                        }
                        else
                            this.LogTask.Information("Erro ao carregar a pessoa.");
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
            StreamReader reader = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\DeleteCPFNulo.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.LogTask.Information("Conectando com o BIS.");
                if (!this.BISManager.Connect())
                    throw new Exception("Erro ao conectar com o BIS.");

                this.LogTask.Information("BIS conectado com sucesso!");

                reader = new StreamReader(@"c:\temp\SG3_CPF_Nulo.csv");
                string line = null;
                int totallines = 0;
                int currentline = 1;
                while ((reader.ReadLine() != null))
                    totallines++;
                reader.Close();
                reader = new StreamReader(@"c:\temp\SG3_CPF_Nulo.csv");

                while ((line = reader.ReadLine()) != null)
                {
                    string[] arr = line.Split(';');
                    string persid = arr[0];
                    string nome = arr[1];

                    this.LogTask.Information(String.Format("{0} de {1}. Excluindo: Persid: {2}, Nome: {3}", (currentline++).ToString(), totallines.ToString(),
                        persid, nome));

                    string sql = String.Format("select distinct per.persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, '') from bsuser.persons per where persid = '{0}' and persclass = 'E' " +
                        " and status = 1 and persclassid = '0013B4437A2455E1'", persid);
                    using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                    {
                        if (table != null && table.Rows.Count > 0)
                        {
                            this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                            int count = 1;
                            foreach (DataRow row in table.Rows)
                            {
                                this.LogTask.Information("Excluindo a pessoa.");
                                using (BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection))
                                {
                                    if (persons.DeletePerson(row["persid"].ToString()) == BS_ADD.SUCCESS)
                                        this.LogTask.Information("Pessoa excluida com sucesso.");
                                    else
                                        this.LogTask.Information("Erro ao excluir a pessoa.");
                                }
                            }
                        }
                    }
                }
                reader.Close();
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

        private async void button19_Click(object sender, EventArgs e)
        {
            List<SG3Info> sg3Info = new List<SG3Info>();
            StreamWriter writer = new StreamWriter(@"c:\temp\checksg3cartao.csv");
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\BlockSG3TaskLog.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                this.LogResult = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\SG3\\BlockSG3TaskLog.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                int total = 0;
                foreach (SG3Host host in this.SG3Hosts)
                {
                    this.LogTask.Information(String.Format("Pesquisando a base: {0}", host.Username));
                    User user = new User() { username = host.Username, password = host.Password };
                    using (var client = new HttpClient())
                    {
                        HttpResponseMessage message = await client.PostAsync(String.Format("{0}/{1}", host.Host, host.LoginCommand), new StringContent(JsonConvert.SerializeObject(user), Encoding.UTF8, "application/json"));
                        string response = await message.Content.ReadAsStringAsync();
                        user = (User)JsonConvert.DeserializeObject<User>(response);

                        var httpClient = new HttpClient();

                        httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", user.token);

                        // Send the request and get the response
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
                                    writer.WriteLine(String.Format("{0};{1};{2};{3};{4}", l.cpf, l.colaborador, l.empresa_terceira, l.estabelecimento, l.cracha));
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
                                        writer.WriteLine(String.Format("{0};{1};{2};{3};{4}", l.cpf, l.colaborador, l.empresa_terceira, l.estabelecimento, l.cracha));
                                    }
                                    else
                                    {
                                        this.LogTask.Information("Excluindo....");
                                        writer.WriteLine(String.Format("{0};{1};{2};{3};{4}", l.cpf, l.colaborador, l.empresa_terceira, l.estabelecimento, l.cracha));
                                    }
                                }
                            }
                        }
                    }
                }
                writer.Close();
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Rotina finalizada..");
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            StreamWriter writer = null;
            try
            {
                this.LogTask = new LoggerConfiguration()
                    .WriteTo.File("c:\\Horizon\\Log\\Femsa\\Import\\AjusteCPF.log", rollingInterval: RollingInterval.Day, rollOnFileSizeLimit: true, fileSizeLimitBytes: 10000000, retainedFileCountLimit: 7)
                    .CreateLogger();

                writer = new StreamWriter(@"c:\temp\CheckCPF.csv");
                string line = null;
                BSPersonsInfoFEMSA personsFemsa = null;
                int totallines = 0;
                int currentline = 1;
                string sql = "select per.persid, RE = persno, unicode(lastname), Nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), cardno, status = cd.status, display = cla.DISPLAYTEXTCUSTOMER, Empresa = cmp.NAME, Unidade = cli.NAME from bsuser.persons per " +
                    "left join bsuser.cards cd on cd.persid = per.persid inner join bsuser.PERSCLASSES cla on cla.PERSCLASSID = per.PERSCLASSID inner join bsuser.clients cli on cli.CLIENTID = per.CLIENTID " +
                    "left outer join bsuser.COMPANIES cmp on cmp.COMPANYID = per.COMPANYID where persclass = 'E' and per.status = 1 and(grade != 'SG3' or grade is null)";
                using (DataTable table = this.bisACEConnection.loadDataTable(sql))
                {
                    if (table != null && table.Rows.Count > 0)
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5};{6};{7}", row["persid"].ToString(), row["re"].ToString(), row["nome"].ToString(), row["cardno"].ToString(), row["status"].ToString(), row["display"].ToString(), row["empresa"].ToString(), row["unidade"].ToString()));
                        }
                    }
                }
                //reader.Close();
                writer.Close();
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
            }
            finally
            {
                this.LogTask.Information("Rotina finalizada.");
            }
        }
    }
}
