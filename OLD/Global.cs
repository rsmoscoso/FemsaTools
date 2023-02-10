using HzBISCommands;
using HzLibConnection.Data;
using Serilog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FemsaTools
{
    public class Global
    {
        /// <summary>
        /// Retorna a string de acordo com o comprimento desejado.
        /// </summary>
        /// <param name="text">String desejada.</param>
        /// <param name="length">Comprimento desejado.</param>
        /// <returns></returns>
        public static string getString(string text, int length)
        {
            return text.Length > length ? text.Substring(0, length) : text;
        }

        /// <summary>
        /// Retorna coleção com o cabeçalho da importação de acordo
        /// com o separador selecionado.
        /// </summary>
        /// <param name="line">Primeira linha do arquivo.</param>
        /// <returns></returns>
        public static Dictionary<string, int> getReaders(string line)
        {
            try
            {
                Dictionary<string, int> retval = new Dictionary<string, int>();
                string[] headers = line.Split(',');
                int i = 0;
                foreach (string head in headers)
                {
                    retval.Add(head, i++);
                }

                return retval;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static List<BSPersonsInfo> readCSV(OpenFileDialog fileDialog)
        {
            List<BSPersonsInfo> retval = null;
            StreamReader reader = null;
            Dictionary<string, int> headers = new Dictionary<string, int>();
            string line = null;
            int lineCount = 0;
            int propertiesCount = 0;
            try
            {
                if (fileDialog.ShowDialog() != DialogResult.OK)
                    return retval;

                reader = new StreamReader(fileDialog.FileName);
                while ((line = reader.ReadLine()) != null)
                {
                    if (lineCount++ == 0)
                    {
                        headers = Global.getReaders(line);
                        continue;
                    }

                    string[] data = line.Split(',');
                    BSPersonsInfo personsInfo = new BSPersonsInfo();
                    propertiesCount = 0;

                    for (int i = 0; i < headers.Count; ++i)
                    {
                        string header = headers.ElementAt(i).Key.ToUpper();
                        if (((BSPersonsInfo)personsInfo).GetType().GetProperties().Any(prop => prop.Name.ToUpper().Equals(header)))
                        {
                            if (((BSPersonsInfo)personsInfo).GetType().GetProperty(header).PropertyType == Type.GetType("System.DateTime"))
                            {
                                DateTime dt = DateTime.Now;
                                DateTime.TryParse(data[i], System.Globalization.CultureInfo.GetCultureInfo("pt-BR"),
                                    System.Globalization.DateTimeStyles.None,
                                    out dt);
                                ((BSPersonsInfo)personsInfo).GetType().GetProperty(header).SetValue(personsInfo, dt);
                            }
                            else
                                ((BSPersonsInfo)personsInfo).GetType().GetProperty(header).SetValue(personsInfo, data[i], null);
                        }
                        else
                        {
                            // Verifica o cabeçalho das autorizações, cartões
                            if (!String.IsNullOrEmpty(data[i]) && header.IndexOf("AUTORIZACAO") > -1)
                            {
                                BSAuthorizationInfo binfo = new BSAuthorizationInfo();
                                binfo.SHORTNAME = data[i];
                                if (personsInfo.AUTHORIZATIONS.Find(auth => auth.SHORTNAME.ToLower().Equals(data[i].ToLower())) == null)
                                    personsInfo.AUTHORIZATIONS.Add(binfo);
                            }
                        }
                    }

                    if (retval == null)
                        retval = new List<BSPersonsInfo>();
                    retval.Add(personsInfo);
                }

                reader.Close();

                return retval;
            }
            catch (Exception ex)
            {
                if (reader != null)
                    reader.Close();

                return retval;
            }
        }

        #region Serialize
        /// <summary>
        /// Serialização de uma tabela.
        /// </summary>
        /// <typeparam name="T">Classe para serialização.</typeparam>
        /// <param name="table">Tabela a ser serializada.</param>
        /// <returns></returns>
        public static List<T> ConvertDataTable<T>(DataTable table)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in table.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }

            return data;
        }

        /// <summary>
        /// Retorna um item da tabela.
        /// </summary>
        /// <typeparam name="T">Classe serializada.</typeparam>
        /// <param name="row">Linha da tabela.</param>
        /// <returns></returns>
        public static T GetItem<T>(DataRow row)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();
            object objvalue = "";
            string name = null;
            try
            {
                foreach (DataColumn column in row.Table.Columns)
                {
                    foreach (PropertyInfo pro in temp.GetProperties())
                    {
                        if (pro.Name.ToLower() == column.ColumnName.ToLower())
                        {
                            name = column.ColumnName;

                            //if (name.ToLower().Equals("photoimage"))
                            //{
                            //    if (!String.IsNullOrEmpty(row["photoimage"].ToString()))
                            //    {
                            //        Byte[] data = new byte[0];
                            //        data = (Byte[])(row["photoimage"]);
                            //        objvalue = data;
                            //    }
                            //}

                            if (row[column.ColumnName] is System.DBNull)
                            {
                                if (column.DataType == Type.GetType("System.String"))
                                    objvalue = "";
                                else if (column.DataType == Type.GetType("System.DateTime"))
                                    objvalue = new DateTime(1, 1, 1);
                                else if (column.DataType == Type.GetType("System.Byte[]"))
                                    objvalue = null;
                                else
                                    objvalue = 0;
                            }
                            else
                            {
                                if (column.DataType == Type.GetType("System.String"))
                                    objvalue = row[column.ColumnName].ToString();
                                else if (column.DataType == Type.GetType("System.DateTime"))
                                    objvalue = DateTime.Parse(row[column.ColumnName].ToString());
                                else if (column.DataType == Type.GetType("System.Int32"))
                                    objvalue = int.Parse(row[column.ColumnName].ToString());
                                else if (column.DataType == Type.GetType("System.Decimal"))
                                    objvalue = decimal.Parse(row[column.ColumnName].ToString());
                                ;

                            }
                            pro.SetValue(obj, objvalue, null);
                        }
                        else
                            continue;
                    }
                }
                return obj;
            }
            catch (Exception ex)
            {
                string x = ex.Message;
                return obj;
            }



        }
        #endregion

    }
}
