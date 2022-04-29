using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Tools.Excel;

using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;

namespace MyDataWorkd
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 


    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

      private void Removedups()
        {
            if (ReportSelect.Text == "Report A")
            {
                Logic lc = new Logic();
                lc.KeyA = KeyA.Text;
                lc.NumericA = NumericA.Text;
                lc.RemoveDups(Data.TableA);
                UpdateTableA();
            } else if (ReportSelect.Text == "Report B")
            {
                Logic lc = new Logic();
                lc.KeyA = KeyB.Text;
                lc.NumericA = Numeric2.Text;
                lc.RemoveDups(Data.TableB);
                UpdateTableB();
            }

        }


        private void ImpExcel()
        {
            Logic lg = new Logic();
            lg.SheetUrl = BrowseTextBox.Text.ToString();
            lg.SheetName = SheetName.Text.ToString();
            lg.ReportName = ReportSelect.Text.ToString();
            

            if (ReportSelect.Text == "Report A")
            {
                ListBoxA.Items.Clear();
               
                lg.FillDataGrids(ReportA, ListBoxA, KeyA, NumericA);

            }
            else if (ReportSelect.Text == "Report B")
            {
              
                ListBoxB.Items.Clear();
                lg.FillDataGrids(ReportB, ListBoxB, KeyB, Numeric2);

            }

        }

        private void Imp2()
        {
            Logic lg = new Logic();
            Functions fn = new Functions();
            fn.updateData(ListBoxA, Data.TableA, ReportA);
            fn.updateData(ListBoxB, Data.TableB, ReportB);
        }
        //private void removeColumn()
        //{

        //    Functions fn = new Functions();
        //    //Remove function
        //    if (Action.Text.ToString() == "Remove")
        //    {
        //        if (ReportSelect.Text.ToString() == "Report A")
        //        {
        //            fn.RemoveColumns(ListBoxA, Data.TableA);
        //            fn.updateData(ListBoxA, Data.TableA, ReportA);
                  
        //    }
        //    else if (ReportSelect.Text.ToString() == "Report B")
        //    {
        //        fn.RemoveColumns(ListBoxB, Data.TableB);
        //            fn.updateData(ListBoxB, Data.TableB, ReportB);
        //    }

        //}

        //    if (Action.Text.ToString() == "Join Column")
        //    {
        //        addCL();
        //    }

        //    if (Action.Text.ToString() == "Reconcile")
        //    {
        //        reconciliation();
        //    }

        //}
        
        private void ApplyFuncions()
        {
            Functions fn = new Functions();
            //Remove function
            if (Action.Text.ToString() == "Remove Column")
            {
               MessageBoxResult result = MessageBox.Show("Are you sure you want to permenantly remove selected column from Report A and Report B ?", "Warning"
                    , MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {

                    fn.RemoveColumns(ListBoxA, Data.TableA);
                    fn.updateData(ListBoxA, Data.TableA, ReportA);
                    
                    fn.RemoveColumns(ListBoxB, Data.TableB);
                    fn.updateData(ListBoxB, Data.TableB, ReportB);
                }

                

            }

            if (Action.Text.ToString() == "Join Column")
            {
                addCL();
            }

            if (Action.Text.ToString() == "Reconcile")
            {
                reconciliation();
            }

            if (Action.Text.ToString() == "Combine Duplicates")
            {
                Removedups(); 
            }

            if (Action.Text.ToString() == "Generate Email")
            {
                CreateEmail();
            }

        }

        private void Export()
        {
            


        }
        private void addCL()
        {
            Logic lg = new Logic();
            lg.NumericA = NumericA.Text.ToString();
            lg.NumericB = Numeric2.Text.ToString();
            lg.KeyA = KeyA.Text.ToString();
            lg.KeyB = KeyB.Text.ToString();

            //lg.AddColumn(ListBoxB,ListBoxA);
            lg.AddColumn2(ListBoxB, ListBoxA);

            Functions fn = new Functions();
            fn.updateData(ListBoxA, Data.TableA, ReportA);
        }
        private void Load1_Click(object sender, RoutedEventArgs e)
        {
            //Imp2();

          
            if (ReportSelect.Text == "")
            {
                MessageBox.Show("Please select the report you want to import and try again");
            }
            else
            {

                ImpExcel();

            }

        }

  

        private void BrowseFile()
        {
            Functions fn = new Functions();
            fn.UrlTextbox = BrowseTextBox.Text.ToString();
            fn.Openfile(BrowseTextBox);
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            BrowseFile();
        }

        private void reconciliation()
        {
            Logic lg = new Logic();
            lg.NumericA = NumericA.Text.ToString();
            lg.NumericB = Numeric2.Text.ToString();
            lg.KeyA = KeyA.Text.ToString();
            lg.KeyB = KeyB.Text.ToString();
            lg.reconcile();
            Functions fun = new Functions();
            fun.updateData(ListBoxA, Data.TableA, ReportA);
            fun.updateData(ListBoxB, Data.TableB, ReportB);
        }

    
        // Emails related fucntions >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        private void EnableEmail()
        {
            String EmailCheck ="";
            foreach (DataColumn Dc in Data.TableA.Columns)
            {
                if (Dc.ColumnName == "Send Email")
                {
                    EmailCheck = "Yes";
                }
              

            }

            if (EmailCheck != "Yes")
            {
                var column = new DataColumn("Send Email", typeof(bool));
                column.DefaultValue = false;

                Data.TableA.Columns.Add(column);
            } else
            {
                MessageBox.Show("It looks like Email function was already activated");
            }

            UpdateTableA();
        }

        private void UpdateTableA()
        {

            Functions fun = new Functions();
            fun.updateData(ListBoxA, Data.TableA, ReportA);
        }
        private void UpdateTableB()
        {
            Functions fn = new Functions();
            fn.updateData(ListBoxB, Data.TableB, ReportB);
        }
        private void EditEmailTemplate()
        {
            EmailG eg = new EmailG();
            eg.Show();

            UpdateTableA();

        }

        private void CreateEmail()
        {
            EmailBody Eb = new EmailBody();
            Eb.CreatEmail(Data.TableA);
        }
      
// End of Emails Related Functions >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
     

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            //removeColumn();
            ApplyFuncions();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            EnableEmail();
        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            BrowseFile();
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            ImpExcel();
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            EditEmailTemplate();
        }

        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            My_DataTable_Extensions.ExportToExcel(Data.TableA, "");
        }
    }
}
