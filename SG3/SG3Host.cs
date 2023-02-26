using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools.SG3
{
    public class SG3Host
    {
        #region Variables
        /// <summary>
        /// Host do SG3.
        /// </summary>
        public string Host { get; set; }
        /// <summary>
        /// Endpoint do SG3.
        /// </summary>
        public string Command { get; set; }
        /// <summary>
        /// Usuario.
        /// </summary>
        public string Username { get; set; }
        /// <summary>
        /// Senha.
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// Comando para o token de login.
        /// </summary>
        public string LoginCommand { get; set; }

        #endregion
    }
}
