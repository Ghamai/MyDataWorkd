using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;

namespace MyDataWorkd
{
    class Connection
    {
        public string Url { get; set; }
        public string Sheet { get; set; }
        public string Extension { get; set; }

        public DataTable tableMain = new DataTable();
        string pathcon;

        public DataTable Table1()
        {

            // this table will return a table with data

            if (Extension.CompareTo(".xls") == 0)
            {

                 pathcon = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Url + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";
            }
            else
            {
                pathcon = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Url + ";Extended Properties='Excel 12.0;HDR=1';"; //for above excel 2007  
            }
            OleDbConnection connect = new OleDbConnection(pathcon);
                //OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from[Member Upload$]", connect);
                OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from [Sheet1$]", connect);
                DataTable dt = new DataTable();
                datadap.Fill(dt);
                connect.Close();
            tableMain = dt.Copy();
        
            return dt;
        }
        public DataTable Table2()
        {

            // this table will return a table with data


            //string pathcon = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Url + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";


            if (Extension.CompareTo(".xls") == 0)
            {

                pathcon = "provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Url + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";
            }
            else
            {
                pathcon = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Url + ";Extended Properties='Excel 12.0;HDR=1';"; //for above excel 2007  
            }

            OleDbConnection connect = new OleDbConnection(pathcon);
            //OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from[Member Upload$]", connect);
            OleDbDataAdapter datadap = new OleDbDataAdapter("Select*from [Sheet1$]", connect);
            DataTable dt = new DataTable();
            datadap.Fill(dt);
            connect.Close();


            return dt;
        }


        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=1';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + Sheet + "$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  

                    for (int i = dtexcel.Rows.Count - 1; i >= 0; i--)
                    {
                        if (dtexcel.Rows[i][1] == DBNull.Value)
                            dtexcel.Rows[i].Delete();
                    }
                    dtexcel.AcceptChanges();
                }
                catch { }
            }
            return dtexcel;
        }


    }
}
