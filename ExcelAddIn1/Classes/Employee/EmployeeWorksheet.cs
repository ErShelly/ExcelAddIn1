using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Data;
using System.Windows.Forms;
using System.Reflection;
using System.ComponentModel;

namespace ExcelAddIn1.Employee.Classes
{

    public class EmployeeWorksheet
    {

        public static void CreateEmployeeTemplateWorksheet()
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.GetExcelApplication();

                Excel.Worksheet newWorksheet;
                newWorksheet = (Excel.Worksheet)app.ActiveWorkbook.Worksheets.Add();

                try
                {
                    newWorksheet.Name = "Employee";
                }
                catch
                {
                    
                }

                var emp = new Employee();
                
                int col = 1;
                int count = emp.count();

                Type type = emp.GetType();
                PropertyInfo[] properties = type.GetProperties();

                foreach (PropertyInfo property in properties)
                {
                    Excel.Range columnRange = app.get_Range(col);
                    columnRange.Value = property.Name;

                    AttributeCollection attributes = TypeDescriptor.GetProperties(property)[property.Name].Attributes;
                    DescriptionAttribute myAttribute = (DescriptionAttribute)attributes[typeof(DescriptionAttribute)];
                    columnRange.Value2 = myAttribute.Description;
                    col++;
                }

            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message,"Error", MessageBoxButtons.OK);
            }
            return;
        }

    }
}
