using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using OUtlookApp = Microsoft.Office.Interop.Outlook.Application;

namespace MyDataWorkd
{
   public class EmailBody
    {
    
      public void CreatEmail(System.Data.DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                try
                {
                    if (row["Send Email"] is true)
                    {




                        OUtlookApp outlookApp = new OUtlookApp();
                        MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);
                        mailItem.To = row[Data.ETo].ToString();
                        mailItem.Subject = Data.Esubject1 +" "+ row[Data.Esubject].ToString();
                        mailItem.Body = Data.EBody1 + " " + row[Data.ECombo1] + " " + Data.EBody2 + " " + row[Data.ECombo2] + " " + Data.EBody3;
                        mailItem.Display(false);
                    }

                }catch(System.ArgumentNullException)
                {
                    System.Windows.MessageBox.Show("Please check you email template and try again");
                }
                //catch (System.Exception)
                //{
                //    System.Windows.MessageBox.Show("Somthing went wrong please try again");
                //}

            }
        }
        

    }
}
