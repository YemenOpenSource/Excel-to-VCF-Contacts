using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;

namespace excel2vcard.Models
{
    public class send
    {
        public string name { get; set; }
        public string value { get; set; }
    }
    public class returned
    {
        public List<int> ErrorList = new List<int>();
        public String messageType;
    }

    public  static class messageTypes
    {
        public const string Dublicate = "Dublicate";
        public const string Done = "Done";
        public const string Failed = "Failed";

        public const string NotValiedColumons = "NotValiedColumons";
        public const string UnknownError = "Unknown Error";
        public const string UnknownErrorline = "Unknown Error in Line";

        public const string NotValiedLocation = "Not Valied Location Data";        
        public const string NotValiedSwich = "Not Valied switch Data";
        public const string NotValiedtypeOfservce = "Not Valied type of Servce";

        public const string NotValiedGovernorate = "Not Valied Governorate Data";
        public const string NotValiedBranch = "Not Valied Branch Data";
        public const string NotValiedAbilty = "Data of abilty is incorrect";


    }


    public class Reader
    {


        public DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            DataTable dt = new DataTable();
            using (StreamReader sr = new StreamReader(strFilePath))
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    if (rows.Length > 1)
                    {
                        DataRow dr = dt.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i].Trim();
                        }
                        dt.Rows.Add(dr);
                    }
                }

            }


            return dt;
        }

        public DataTable ConvertXSLXtoDataTable(string strFilePath, string connString)
        {
            OleDbConnection oledbConn = new OleDbConnection(connString);
            DataTable dt = new DataTable();
            try
            {

                oledbConn.Open();
                using (OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn))
                {
                    OleDbDataAdapter oleda = new OleDbDataAdapter();
                    oleda.SelectCommand = cmd;
                    DataSet ds = new DataSet();
                    oleda.Fill(ds);

                    dt = ds.Tables[0];
                }
            }
            catch
            {
            }
            finally
            {

                oledbConn.Close();
            }

            return dt;

        }


        
    }
}