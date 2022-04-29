using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;




namespace MyDataWorkd
{
    class Functions
    {
        public string UrlTextbox { get; set; }



        public void Openfile(TextBox textBox)
        {
           
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
                Nullable<bool> result = openFileDialog.ShowDialog();
                if (result == true)
                {
                    textBox.Text = openFileDialog.FileName;

                }

               
            
        }
        public void updateData(ListBox listBox,DataTable dataTable,DataGrid dataGrid)
        {
            listBox.Items.Clear();
            foreach (DataColumn cl in dataTable.Columns)
            {
                listBox.Items.Add(cl.ColumnName.ToString());
            }
            dataGrid.ItemsSource = null;
            dataGrid.ItemsSource = dataTable.DefaultView;
        }

       
        public void RemoveColumns(ListBox listBox,DataTable dataTable)
        {
            foreach (string s in listBox.SelectedItems)
            {
                dataTable.Columns.Remove(s);

            }
            

        }

        
        
    }
}
