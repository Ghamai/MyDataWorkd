using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.IO;
using System.Windows.Media;

namespace MyDataWorkd
{
    class Logic
    {

        public string SheetName { get; set; }
        public string SheetUrl { get; set; }
        public string KeyA { get; set; }
        public string KeyB { get; set; }
        public string NumericA { get; set; }
        public string NumericB { get; set; }
        public string ReportName { get; set; }
        public string AddedColumnName { get; set; }

        public DataTable table1 = new DataTable();
        public static DataTable table2 = new DataTable();
        public void FillDataGrids(DataGrid dataGrid, ListBox listBox, ComboBox comboBox1, ComboBox comboBox2)
        {
            Connection cn = new Connection();
            cn.Sheet = SheetName;
            cn.Url = SheetUrl;
            cn.Extension = Path.GetExtension(SheetUrl);
            cn.Table1();
            // to keep the table which will be manipulated further separate and connected to Datagrid from begining this if statment is used.
            try
            {
                if (ReportName == "Report A")
                {
                    Data.TableA = cn.tableMain.Copy();
                    dataGrid.ItemsSource = Data.TableA.DefaultView;
                    foreach (DataColumn cl in Data.TableA.Columns)
                    {
                        listBox.Items.Add(cl.ColumnName.ToString());
                        comboBox1.Items.Add(cl.ColumnName.ToString());
                        comboBox2.Items.Add(cl.ColumnName.ToString());


                    }
                }
                else if (ReportName == "Report B")
                {

                    Data.TableB = cn.tableMain.Copy();
                    dataGrid.ItemsSource = Data.TableB.DefaultView;
                    foreach (DataColumn cl in Data.TableB.Columns)
                    {
                        listBox.Items.Add(cl.ColumnName.ToString());
                        comboBox1.Items.Add(cl.ColumnName.ToString());
                        comboBox2.Items.Add(cl.ColumnName.ToString());

                    }

                }
            }
            catch (Exception)
            {
                MessageBox.Show("Somthing went wrong Please check sheet names");
            }
        }


        public void RColumns(ListBox listBox)
        {
            foreach (string s in listBox.SelectedItems)
            {
                MessageBox.Show(s);
            }
        }


        public void reconcile()
        {

            foreach (DataColumn cl in Data.TableA.Columns)
            {
                // Check if the clomn already exist
                if (cl.ColumnName == ("Reconciliation"))
                {
                    Data.Reconcile = "Created";
                }
            }
            foreach (DataColumn cl in Data.TableB.Columns)
            {// Check if the clomn already exist
                if (cl.ColumnName == ("Not Found"))
                {
                    Data.Reconcile = "Created";
                }
            }

            if (Data.Reconcile != "Created")
            {
                Data.TableA.Columns.Add("Reconciliation");

            }
            else
            {
                MessageBox.Show("Reconcilliation Column already exists system will use the same column");
            }
            if (Data.Reconcile != "Created")
            {
                Data.TableB.Columns.Add("Not Found");

            }
            else
            {
                MessageBox.Show("Reconcilliation Column already exists system will use the same column");
            }



            foreach (DataRow row in Data.TableA.Rows)
            {
                foreach (DataRow row2 in Data.TableB.Rows)
                {
                    if (row[KeyA].ToString() == row2[KeyB].ToString() && row[NumericA].ToString() != row2[NumericB].ToString())
                    {
                        row["Reconciliation"] = "Not Matching";
                    }
                    if (row2[KeyB].ToString() == row[KeyA].ToString())
                    {
                        row2["Not Found"] = "Found in A";
                    }


                }

            }

        }


        public void AddColumn2(ListBox listBoxB, ListBox listbaxA)
        {
            //last time I duplicated added column method to add a unique for each duplicate colum
            foreach (string CName in listBoxB.SelectedItems)
            {
                int ColCoun = 0;
                Data.AddedColumns.Add(CName);
                foreach (string existingColumn in Data.AddedColumns)
                {
                    if (CName == existingColumn)
                    {
                        ColCoun++;
                    }
                }
                Data.TableA.Columns.Add(CName + "RB" + ColCoun.ToString());

                foreach (DataRow row in Data.TableA.Rows)
                {
                    foreach (DataRow row2 in Data.TableB.Rows)
                    {
                        if (row[KeyA].ToString() == row2[KeyB].ToString())
                        {
                            row[CName + "RB" + ColCoun.ToString()] = row2[CName].ToString();
                        }
                    }
                }
            }
        }

        public void RemoveDups(DataTable dataTable)

        {
            for (int i = 0; i < dataTable.Rows.Count - 1; i++) //compare data
            {
                var Row = dataTable.Rows[i];
                string abc = Row[KeyA].ToString(); ; //+ Row.Cells[1].Value.ToString().ToUpper();

                for (int j = i + 1; j < dataTable.Rows.Count; j++)
                {
                    try
                    {

                        var Row2 = dataTable.Rows[j];
                        string def = Row2[KeyA].ToString(); //+ Row2.Cells[1].Value.ToString().ToUpper();
                        if (abc == def)
                        {
                            decimal kk;
                            decimal kk2;
                            if (decimal.TryParse(Row[NumericA].ToString(), out kk) && decimal.TryParse(Row2[NumericA].ToString(), out kk2))
                            {
                                Row[NumericA] = Convert.ToDecimal(Row[NumericA]) + Convert.ToDecimal(Row2[NumericA]);
                                dataTable.Rows.Remove(Row2);
                                j--;
                            }
                            else
                            {
                                Data.TableA.Columns.Add("Errors");
                                Row["Errors"] = "Cant Convert to Number";
                                Row2
                                    ["Errors"] = "Cant Convert to Number";
                                MessageBox.Show("Error in numbers conversion");
                            }
                        }
                    }
                    catch(Exception)
                    {
                        MessageBox.Show("Somthing went wrong please try again");
                    }
                }
            }



        }
    }
}
