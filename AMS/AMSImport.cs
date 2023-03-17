using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using System.Xml.Linq;
using System.IO;

namespace FemsaTools.AMS
{
    public class AMSImport
    {
        #region Variables
        /// <summary>
        /// Log da aplicação.
        /// </summary>
        public ILogger LogTask { get; set; }
        #endregion

        public StreamWriter parseLine(StreamReader reader, StreamWriter writer, string line)
        {
            try
            {
                AMSEvent retval = new AMSEvent();
                DateTime dateTime = DateTime.Now;
                int pos = 0;

                string[] arr = line.Split(';');
                bool isDateTime = DateTime.TryParse(arr[0], out dateTime);
                if (isDateTime)
                {
                    if ((pos = arr[2].IndexOf("via")) > 0)
                    {
                        retval.Data = dateTime;
                        retval.Door = arr[2].Substring(pos + 4, arr[2].Length - (pos + 4));
                        retval.Cardno = arr[3];
                        line = reader.ReadLine();
                        arr = line.Split(';');
                        retval.Name = String.Format("{0} {1}", arr[1], arr[0]);
                        retval.Persno = arr[3];
                        line = reader.ReadLine();
                        arr = line.Split(';');
                        isDateTime = DateTime.TryParse(arr[0], out dateTime);
                        if (!isDateTime)
                        {
                            retval.Company = arr[0];
                            this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5}", retval.Data.ToString("dd/MM/yyyy HH:mm:ss"), retval.Persno, retval.Name, retval.Cardno, retval.Company, retval.Door));
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", retval.Data.ToString("dd/MM/yyyy HH:mm:ss"), retval.Persno, retval.Name, retval.Cardno, retval.Company, retval.Door));
                        }
                        else
                        {
                            retval.Company = "";
                            this.LogTask.Information(String.Format("{0};{1};{2};{3};{4};{5}", retval.Data.ToString("dd/MM/yyyy HH:mm:ss"), retval.Persno, retval.Name, retval.Cardno, retval.Company, retval.Door));
                            writer.WriteLine(String.Format("{0};{1};{2};{3};{4};{5}", retval.Data.ToString("dd/MM/yyyy HH:mm:ss"), retval.Persno, retval.Name, retval.Cardno, retval.Company, retval.Door));

                            this.parseLine(reader, writer, line);
                        }
                    }
                }

                return writer;
            }
            catch (Exception ex)
            {
                return writer;
            }
        }

        #region Events
        /// <summary>
        /// Construtor da classe.
        /// </summary>
        /// <param name="logger">Log da aplicação.</param>
        public AMSImport(ILogger logger)
        {
            this.LogTask = logger;
        }
        #endregion
    }
}
