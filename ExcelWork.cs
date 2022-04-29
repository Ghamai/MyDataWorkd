using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace MyDataWorkd
{
    public class ExcelWork
    {

        public void ExTry()
        {
            //Worksheet current = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Application excelAPP = new Excel.Application();
            ////Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add();
            //Excel.Worksheet newWorksheet;
            //newWorksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            //newWorksheet.Name = "Report";
            excelAPP.ActiveWorkbook.Worksheets.Add();
        }
    }
}
