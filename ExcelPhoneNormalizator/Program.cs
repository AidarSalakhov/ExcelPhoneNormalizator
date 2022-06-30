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
                    if (helper.Open(Path.Combine(Environment.CurrentDirectory, "messages.csv")))
                    {
                        helper.SaveAsTXT(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                        helper.Dispose();

                        helper.Open(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                        helper.DeleteColumn("B1:X1");

                        helper.removeDuplicatesA();

                        Console.WriteLine("Нормализация телефонов...");

                        helper.Normalize();

                        helper.removeDuplicatesB();

                        helper.DeleteColumn("A1");

                        helper.DeleteEntireRow("A1");

                        helper.DeleteColumn("B1");

                        helper.SetColumnWidth(1, 18);

                        helper.SaveAsXLSX(Path.Combine(Environment.CurrentDirectory, $"leads-count-{helper.LastRealRow()}.xlsx"));

                        File.Delete(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                        File.Delete(Path.Combine(Environment.CurrentDirectory, "messages.csv"));

                        Console.Clear();

                        Console.WriteLine($"Количество чистых заявок: {helper.LastRealRow()}");

                        helper.Dispose();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
