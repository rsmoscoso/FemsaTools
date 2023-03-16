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
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools.SG3
{
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

        /// <summary>
        /// Importa os dados do SG3.
        /// </summary>
        /// <param name="liberacao">Classe com os dados do SG3.</param>
        public void Import(Liberacao liberacao)
        {
            try
            {
                string firstname = "";
                string lastname = "";
                int i = 1;
                BSCommon.ParseName(liberacao.colaborador.Replace("'", ""), out firstname, out lastname);

                string sql = String.Format("select persid, persno, nome = isnull(firstname, '') + ' ' + isnull(lastname, ''), persclassid from bsuser.persons where status = 1 and persclass = 'E' and persno = '{0}'",
                    liberacao.cpf.Replace("'", ""));
                using (DataTable table = this.bisConnection.loadDataTable(sql))
                {
                    if (table.Rows.Count > 0)
                    {
                        this.LogTask.Information(String.Format("{0} registros encontrados.", table.Rows.Count.ToString()));
                    }
                    else
                    {
                        this.LogTask.Information("Nova pessoa.");

                        string clientid = this.getClientidFromArea(liberacao.estabelecimento);
                        if (!String.IsNullOrEmpty(clientid))
                        {
                            this.LogTask.Information(String.Format("Alterando a divisao para o ID: {0}", clientid));
                            if (this.bisManager.SetClientID(clientid))
                                this.LogTask.Information("ClientID alterado com sucesso.");
                            else
                                this.LogTask.Information(String.Format("Erro ao alterar o ClientID. {0}", this.bisManager.GetErrorMessage()));
                        }
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
                                this.LogTask.Information("Pessoa gravada com sucesso.");
                            else
                                this.LogTask.Information("Erro ao gravar a pessoa.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                //this.sendErrorMail(this.bisConnection, ex.Message);
            }
        }

        /// <summary>
        /// Deleta os dados do SG3.
        /// </summary>
        /// <param name="liberacao">Classe com os dados do SG3.</param>
        public void Delete(Liberacao liberacao)
        {
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
                                this.LogTask.Information("Pessoa excluida com sucesso.");
                            else
                                this.LogTask.Information("Erro ao excluir a pessoa.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.LogTask.Information(String.Format("Erro: {0}", ex.Message));
                //this.sendErrorMail(this.bisConnection, ex.Message);
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
        public SG3Import(ILogger logtask, ILogger logresult, SQLParameters bissql, SQLParameters bislogsql, BISManager bismanager)
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
