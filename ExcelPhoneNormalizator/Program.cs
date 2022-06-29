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
            Console.WriteLine("Запуск программы, подождите...");

            ExcelHelper helper = new ExcelHelper();

            helper.Open(Path.Combine(Environment.CurrentDirectory, "messages.csv"));

            helper.SaveAs(Path.Combine(Environment.CurrentDirectory, $"leads.xlsx"));

            helper.Open(Path.Combine(Environment.CurrentDirectory, "leads.xlsx"));

            helper.DeleteColumn("A1:E1");

            //helper.DeleteDuplicates("A1:B7");

            //helper.DeleteColumn("B1");

            //Console.WriteLine("Нормализация телефонов...");

            //helper.Normalize();

            //helper.removeDuplicatesB();

            //helper.DeleteColumn("A1");

            //helper.DeleteEntireRow("A1");

            //helper.DeleteColumn("B1");

            //helper.DeleteEntireRow("A1");

            helper.SaveAs(Path.Combine(Environment.CurrentDirectory, $"leads-count-{helper.LastRealRow()}.xlsx"));

            Console.Clear();

            Console.WriteLine($"Количество чистых заявок: {helper.LastRealRow()}");

            helper.Dispose();
        }
    }
}
