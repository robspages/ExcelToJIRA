using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace ExcelToSQL
{
    public class ExcelSheet : DataTable 
    {

        public String Name = String.Empty;
        private FileConnection fileConnection; 
        public ExcelSheet () {}

        public ExcelSheet (String sheetName, FileConnection conn)
        {
           Name = sheetName;
           fileConnection = conn; 
        }

        public void getData()
        { 
            using (fileConnection.Connection)
            {
                try
                {
                    OleDbCommand cmd = new OleDbCommand();
                  //  Name = (Name.Substring(Name.Length - 1) == "$" ? Name : Name + "$");//insert the $ if it isn't there already
                    
                    cmd.CommandText = String.Format("Select * from [{0}]",  Name); 
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = fileConnection.Connection;
                    OleDbDataReader reader = cmd.ExecuteReader();

                    this.Load(reader);

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

    }
}
