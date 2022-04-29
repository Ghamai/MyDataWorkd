using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace MyDataWorkd
{
    
    static class Data
    {

        public static string ETo;
        public static string Esubject;
        public static string Esubject1;
        public static string EBody1;
        public static string ECombo1;
        public static string EBody2;
        public static string ECombo2;
        public static string EBody3;


        public static string Reconcile;
        public static string Found;
        public static int ColumnNumber = 1;
        public static DataTable TableA;
        public static DataTable TableB;
        public static string ReportName { get; set; }
        public static DataTable TablOne()
        {

            Connection cn = new Connection();
            

            TableA = cn.Table1();
            return TableA;
        }
        public static DataTable TablTwo()
        {
            Connection cn = new Connection();
            DataTable TableB = new DataTable();

            TableB = cn.Table1();

            return TableB;

        }

        public static List<string> AddedColumns = new List<string>();

    }
}
