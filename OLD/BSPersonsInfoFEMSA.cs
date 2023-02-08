using HzBISCommands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools
{
    public class BSPersonsInfoFEMSA : BSPersonsInfo
    {
        #region Variables
        public string CPF { get; set; }
        public string RG { get; set; }
        public string PISPASEP { get; set; }
        public string DEFICIENCIA { get; set; }
        public string INTERNOEXTERNO { get; set; }
        public string VA { get; set; }
        public string VT { get; set; }
        public string NOMERECEBEDOR { get; set; }
        public string DATAADMISSAO { get; set; }
        public string CODIGOAREA { get; set; }
        public string DESCRICAOAREA { get; set; }
        public string STATUSPROPRIO { get; set; }
        public string CNPJ { get; set; }
        #endregion

        #region Events
        /// <summary>
        /// Construtor da clsse.
        /// </summary>
        public BSPersonsInfoFEMSA() : base()
        {

        }
        #endregion
    }
}
