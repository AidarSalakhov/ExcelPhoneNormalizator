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
                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "ForNormalization.xlsx")))
                    {
                        Console.WriteLine("Идёт обработка файла...");

                        helper.Normalize();

                        helper.Save();

                        helper.Dispose();

                        Console.WriteLine("Готово!");
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
