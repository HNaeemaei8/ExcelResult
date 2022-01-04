using ExcelOutput.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOutput
{
    public class Result
    {
        public static void Results()
        {
            var p = new Person();
            var model = new List<Person>();
            p.Name = "Mohandes Hosein";
            p.Descriotion = "God of Programming";
            model.Add(p);
            try
            {
                string fileName = "UserManager.xlsx";
                Console.WriteLine("Please give a location to save :");
                string customExcelSavingPath = "D:\\PersonExcel.xlsx" + fileName;
                ExcelExport excel = new ExcelExport();
                ConvertData<Person> prs = new ConvertData<Person>();
                excel.GenerateExcel(prs.ConvertToDataTable(model), customExcelSavingPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
