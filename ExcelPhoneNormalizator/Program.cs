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
                Console.WriteLine("Запуск программы, подождите...");

                using (ExcelHelper helper = new ExcelHelper())
                {
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "messages.csv")))
                    {
                        helper.removeDuplicatesA();

                        helper.DeleteColumn("B1:J1");

                        Console.WriteLine("Нормализация телефонов...");
                        helper.Normalize();

                        helper.removeDuplicatesB();

                        helper.DeleteColumn("A1");

                        helper.DeleteEntireRow("A1");

                        helper.Save();

                        helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "messages.csv"));

                        Console.Clear();

                        Console.WriteLine($"Количество чистых заявок: {helper.LastRow()}");

                        helper.Dispose();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
