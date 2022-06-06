using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelPhoneNormalizator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (BO.ExcelHelper helper = new BO.ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "TestLT.xlsx")))
                    {
                        helper.Set(column: "A", row: 1, data: "lksadklsajdkl");
                        var val = helper.Get(column: "A", row: 6);
                        helper.Set(column: "B", row: 1, data: DateTime.Now);

                        helper.Save();
                    }
                }

                Console.Read();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
