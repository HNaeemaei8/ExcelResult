using System.Collections.Generic;
using System.Data;

namespace ExcelOutput
{
    public class ExcelExport
    {
        public  void GenerateExcel(DataTable dataTable, string path)
         {
            DataSet dataSet = new DataSet();

            //dataTable.Rows.Add("1", "Devesh Omar" );
            //dataTable.Rows.Add("2", "Nikhil Vats");
            dataSet.Tables.Add(dataTable);
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name  
                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                // add all the columns  
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                // add all the rows  
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
       
            }

            excelWorkBook.SaveAs(path); 
            excelWorkBook.Close();
            excelApp.Quit();


        }

    }
}
