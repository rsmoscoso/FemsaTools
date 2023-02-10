using AMSLib;
using AMSLib.ACE;
using AMSLib.Manager;
using AMSLib.SQL;
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
        private void button4_Click(object sender, EventArgs e)
        {
            StreamReader reader = null;
            StreamWriter writer = null;
            try
            {
                lstData.View = View.Details;

                this.ReadParameters();
                this.bisACEConnection = new HzConexao(this.BISACESQL.SQLHost, this.BISACESQL.SQLUser, this.BISACESQL.SQLPwd, "acedb", "System.Data.SqlClient");

                if (openFileDialog1.ShowDialog(this) != DialogResult.OK)
                    return;

                this.Cursor = Cursors.WaitCursor;

                lstData.Items.Clear();

                string filename = openFileDialog1.FileName;
                float length = 0;
                reader = new StreamReader(filename);
                while (reader.ReadLine() != null)
                    length++;
                reader.Close();
                reader = new StreamReader(filename);
                string line = null;
                int pos = -1;
                string cardid = null;
                string deviceid = null;
                BSPersonsInfo person = new BSPersonsInfo();
                //writer.WriteLine("Data;PERSNO;Nome;CardNO;Local");
                string[] vw = new string[8];
                ListViewItem item = null;
                string tipo = getInfoType();
                pgBar.Visible = true;
                pgBar.Refresh();
                float lineCount = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    lineCount++;
                    pgBar.Value = (int)Math.Truncate((lineCount / length) * 100);
                    pgBar.Refresh();
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
                //writer.Close();
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

                BSPersons persons = new BSPersons(this.BISManager, this.bisACEConnection);
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

                        this.LogTask.Information(String.Format("Total Antes: {0}, Total Depois: {1}. Alteracao concluida com sucesso.", totalBefore.ToString(), totalAfter.ToString()));
                    }
                    else
                        this.LogTask.Information("Erro ao alterar a divisao.");
                }
                else
                    this.LogTask.Information("Erro ao carregar a pessoa.");

                MessageBox.Show("Alteração concluída com sucesso!");
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
    }
}
