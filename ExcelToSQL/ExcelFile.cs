using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;


namespace ExcelToSQL
{
    public class ExcelFile
    {
        private FileConnection fileConn;

        private List<ExcelSheet> _Sheets = new List<ExcelSheet>();
        public List<ExcelSheet> Sheets
        { 
            get 
            { 
                if (_Sheets.Count == 0)
                {
                    foreach (String sheetName in getSheetNames())
                    {
                        ExcelSheet sheet = new ExcelSheet(sheetName, fileConn);
                        sheet.getData();
                        _Sheets.Add(sheet);
                    }
                }

                return _Sheets;
            } 
        }

        public ExcelFile(String fileName)
        {
            fileConn = new FileConnection(fileName);
        }

        /*
         * returna an array of Sheet names from your excel document
         */ 
        public String[] getSheetNames()
        {
            using (fileConn.Connection)
            {
                DataTable dt = new DataTable();
                try
                {
                    dt = fileConn.Connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    String[] excelSheets = new String[dt.Rows.Count];
                    if (dt != null)
                    {
                        // Add the sheet name to the string array.
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            excelSheets[i] = dt.Rows[i]["TABLE_NAME"].ToString();
                        }
                    }
                    return excelSheets;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public ExcelSheet findSheetByName(String name)
        {
            int index = 0;
            for (int i = 0; i < Sheets.Count; i++)
            {
                if (Sheets[i].Name == name || Sheets[i].Name == name + "$")
                {
                    index = i;
                    break; // I know... bad bad 
                }
            }

            return Sheets[index];
        }
    }
}