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
using System.Windows.Shapes;
using Microsoft.Office.Interop.Outlook;
using OUtlookApp = Microsoft.Office.Interop.Outlook.Application;
using System.Data;


namespace MyDataWorkd
{
    /// <summary>
    /// Interaction logic for EmailG.xaml
    /// </summary>
    public partial class EmailG : Window
    {
        public EmailG()
        {


            InitializeComponent();
            MainWindow mn = new MainWindow();
            if (Data.TableA != null)
            {
                foreach (DataColumn cl in Data.TableA.Columns)
                {
                    To.Items.Add(cl.ColumnName.ToString());
                    Subject.Items.Add(cl.ColumnName.ToString());
                    Ecomb1.Items.Add(cl.ColumnName.ToString());
                    ECombo2.Items.Add(cl.ColumnName.ToString());
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GenerateEmail();
        }

        private void GenerateEmail()
        {
            //OUtlookApp outlookApp = new OUtlookApp();
            //MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);
            //mailItem.Subject = "Lets Test and learn";
            //mailItem.Body = "I can do great";
            //mailItem.Display(false);
            Data.EBody1 = EText1.Text.ToString();
            Data.ETo = To.Text.ToString();
            Data.Esubject = Subject.Text.ToString();
            Data.EBody2 = Etext2.Text.ToString();
            Data.ECombo1 = Ecomb1.Text.ToString();
            Data.ECombo2 = ECombo2.Text.ToString(); 
            Data.EBody3 = EText3.Text.ToString();
            Data.Esubject1 = subject1.Text.ToString();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
           

            foreach (DataColumn cl in Data.TableA.Columns)
            {
                To.Items.Add(cl.ColumnName.ToString());
                Subject.Items.Add(cl.ColumnName.ToString());
               Ecomb1.Items.Add(cl.ColumnName.ToString());
               ECombo2.Items.Add(cl.ColumnName.ToString());





            }
        }
    }
}
