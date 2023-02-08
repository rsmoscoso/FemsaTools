using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace FemsaTools
{
    public class FemsaNetworkCredential : NetworkCredential
    {
        #region Variables
        /// <summary>
        /// Arquivo compartilhado.
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// Pasta compartilhada.
        /// </summary>
        public string SharingFolder { get; set; }
        #endregion

        #region Functions
        /// <summary>
        /// Abre o arquivo do compartilhamento.
        /// </summary>
        /// <returns></returns>
        public StreamReader OpenFile()
        {
            StreamReader retval = null;
            try
            {
                NetworkCredential net = new NetworkCredential(this.UserName, this.Password, this.Domain);
                using (new NetworkConnection(this.SharingFolder, net))
                {
                    FileInfo fileInfo = new FileInfo(this.SharingFolder + "\\" + this.FileName);
                    retval = new StreamReader(this.SharingFolder + "\\" + this.FileName);
                }

                return retval;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Abre o arquivo do compartilhamento.
        /// </summary>
        /// <returns></returns>
        public StreamReader CopyOpenFile(out FileInfo fileInfo)
        {
            StreamReader retval = null;
            try
            {
                NetworkCredential net = new NetworkCredential(this.UserName, this.Password, this.Domain);
                string path = String.Format(@"c:\horizon\files\{0}", String.Format(this.FileName + "_{0}", DateTime.Now.Day.ToString()));
                using (new NetworkConnection(this.SharingFolder, net))
                {
                    fileInfo = new FileInfo(this.SharingFolder + "\\" + this.FileName);
                    
                    if (File.Exists(path))
                    {
                        File.Delete(path);
                    }
                    File.Copy(this.SharingFolder + "\\" + this.FileName, path);
                    if (File.Exists(path))
                        retval = new StreamReader(path);
                    else
                        throw new Exception("Erro ao copiar o arquivo.");
                }

                return retval;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        #endregion

        #region Events
        /// <summary>
        /// Construtor da classe. 
        /// </summary>
        /// <param name="user">Nome do usuário.</param>
        /// <param name="pwd">Senha do usuário.</param>
        /// <param name="domain">Domínio AD.</param>
        /// <param name="sharingfolder">Pasta compartialhda.</param>
        /// <param name="filename">Arquivo a ser lido.</param>
        public FemsaNetworkCredential(string user, string pwd, string domain, string sharingfolder, string filename) 
            : base(user, pwd, domain)
        {
            this.SharingFolder = sharingfolder;
            this.FileName = filename;
        }
        #endregion
    }
}
