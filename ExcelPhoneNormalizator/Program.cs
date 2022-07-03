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

                ExcelHelper helper = new ExcelHelper();

                if (helper.Open(Path.Combine(Environment.CurrentDirectory, "messages.csv")))
                {
                    helper.SaveAsTXT(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                    helper.Dispose();

                    helper.Open(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                    string projectName = Convert.ToString(helper.Get("A", 2));

                    helper.DeleteColumn("B1:X1");

                    helper.RemoveDuplicatesA();

                    Console.WriteLine("Нормализация телефонов...");

                    helper.Normalize();

                    helper.RemoveDuplicatesB();

                    helper.DeleteColumn("A1");

                    helper.DeleteRow("A1");

                    helper.DeleteColumn("B1");

                    helper.SetColumnWidth(1, 18);

                    helper.SaveAsXLSX(Path.Combine(Environment.CurrentDirectory, $"{helper.GetLastRow()}.xlsx"));

                    File.Delete(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                    File.Delete(Path.Combine(Environment.CurrentDirectory, "messages.csv"));

                    Console.Clear();

                    Console.WriteLine($"Проект: {projectName}\nКоличество чистых заявок: {helper.GetLastRow()}");

                    helper.Dispose();

                    Console.ReadLine();
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
