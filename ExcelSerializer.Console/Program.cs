using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSerializer.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            var serializer = new ExcelSerializer();
            serializer.DisplayInExcel(new[]
            {
                new
                {
                    Name = "Raz", Age = "24"
                },
                new
                {
                    Name = "Saba", Age = "24"
                }
            });
        }
    }
}
