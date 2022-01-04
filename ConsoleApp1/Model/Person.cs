using ExcelOutput.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOutput.Model
{
    public class Person
    {
        public string Name { get; set; }
        public string Descriotion { get; set; }

        public static Person SeedData()
        {
            return new Person()
            {
                Name = "Ali",
                Descriotion = "Barsa"
            };
        }

        //public static List<Objects> SeedData()
        //{
        //    return new List<Objects>()
        //    {
        //         new Objects
        //         {
        //              Name="Ali",
        //              Descriotion="Barsa"
        //         },
        //         new Objects
        //         {
        //               Name="hossein",
        //              Descriotion="sql"
        //         }
        //    };
        //}
    }

}

