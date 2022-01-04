using ExcelOutput.Model;
using FastMember;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOutput
{
   public class ConvertData<T>
    {
       public  DataTable ConvertToDataTable(List<T> data)
        {

         
            DataTable table = new DataTable();
            using (var reader = ObjectReader.Create(data))
            {
                table.Load(reader);
            }
            // creating a data table instance and typed it as our incoming model   
            // as I make it generic, if you want, you can make it the model typed you want.  
            //DataTable dataTable = new DataTable(models.GetType().Name);
            //dataTable. = models;
            ////Get all the properties of that model  
            //PropertyInfo[] Props = models.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);

            //// Loop through all the properties              
            //// Adding Column name to our datatable  
            //foreach (PropertyInfo prop in Props)
            //{
            //    //Setting column names as Property names    
            //    dataTable.Columns.Add(prop.Name);
            //}
            //// Adding Row and its value to our dataTable  
            //foreach (object item in models)
            //{
            //    var values = new object[Props.Length];
            //    for (int i = 0; i < Props.Length; i++)
            //    {
            //        //inserting property values to datatable rows    
            //        values[i] = Props[i].GetValue(item, null);
            //    }
            //    // Finally add value to datatable    
            //    dataTable.Rows.Add(values);
            //}
            return table;
        }

       
    }
}
