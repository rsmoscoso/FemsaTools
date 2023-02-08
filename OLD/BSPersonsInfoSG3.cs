using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools
{
    public class BSPersonsInfoSG3 : BSPersonsInfoFEMSA
    {
        #region Variables
        public string STATUSTERCEIRO { get; set; }
        public string TIPOTERCEIRO { get; set; }
        public string LIBERADO { get; set; }
        public string ESTABELECIMENTO { get; set; }

        #endregion

        #region Events
        /// <summary>
        /// Construtor da clsse.
        /// </summary>
        public BSPersonsInfoSG3() : base()
        {

        }
        #endregion
    }
}
