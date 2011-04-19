using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;


namespace ExcelToSQL
{
    public class FileConnection
    {
        // provider string - varies by Excel version 
        private string oleProvider = "Microsoft.ACE.OLEDB.12.0";

        //addtional properties for identifying the excel type and information
        private List<String> _oleExtendedProperties = new List<String>();
        private List<String> oleExtendedProperties
        {
            get
            {
                if (_oleExtendedProperties.Count == 0)
                {
                    _oleExtendedProperties.Add("Excel 12.0"); //sets Excel version we expect
                    _oleExtendedProperties.Add("HDR=YES"); //Sets the first row of data to be the DataRow.Column.Name
                    _oleExtendedProperties.Add("IMEX=1"); // dunno 
                }

                return _oleExtendedProperties;
            }
        }

        private String FileName = String.Empty;

        public void Open()
        {
            this.Connection.Open();
        }

        public void Close()
        {
            this.Connection.Close();
        }

        public ConnectionState State()
        {
            return this.Connection.State;
        }

        // store the file conn as a property so we can reuse it within the class... 
        private OleDbConnection _fConnection = new OleDbConnection();
        public OleDbConnection Connection
        {
            get
            {
                if (String.IsNullOrEmpty(_fConnection.ConnectionString.ToString()))
                {
                    _fConnection = new OleDbConnection(connectionString(FileName));
                }

                if (_fConnection.State == ConnectionState.Closed)
                {
                    _fConnection.Open();
                }
                return _fConnection;
            }
        }

        public FileConnection(String fileName)
        {
            try
            {
                FileName = fileName;
                _fConnection = new OleDbConnection(connectionString(fileName));
            }
            catch (OleDbException oleEx)
            {
                throw oleEx;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private String connectionString(String fileName)
        {
            try
            {
                StringBuilder builder = new StringBuilder();

                builder.Append(String.Format("Provider={0};", oleProvider));
                builder.Append(String.Format("Data Source={0};", fileName));
                if (oleExtendedProperties.Count > 0)
                {
                    builder.Append("Extended Properties=" + "\"");
                    foreach (String prop in oleExtendedProperties)
                    {
                        builder.Append(prop + ";");
                    }
                    builder.Append("\"");
                }
                return builder.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}