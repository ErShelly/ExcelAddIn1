using System;
using System.ComponentModel;

namespace ExcelAddIn1.Employee.Classes
{
    class Employee
    {
        [Description("Employee ID")]
        public string EmployeeId { get; set; }

        [Description("Employee Name")]
        public string Name { get; set; }

        [Description("Date Of Birth")]
        public DateTime DateOfBirth { get; set; }

        public int count()
        {
           return(typeof(Employee).GetProperties().Length);
        }
    }
}