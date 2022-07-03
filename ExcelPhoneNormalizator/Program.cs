using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace ExcelPhoneNormalizator
{
    internal class Program
    {
        static void Main(string[] args)
        {

            try
            {
                Console.WriteLine("Запуск программы, подождите...");

                ExcelOperations helper = new ExcelOperations();

                if (helper.OpenCSV(Path.Combine(Environment.CurrentDirectory, "messages.csv")))
                {
                    helper.SaveAsTXT(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                    helper.Dispose();

                    helper.OpenCSV(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                    string projectName = Convert.ToString(helper.Get("A", 2));

                    helper.DeleteColumn("B1:X1");

                    helper.RemoveDuplicatesFromColumn("A");

                    Stopwatch stopWatch = new Stopwatch();

                    stopWatch.Start();

                    helper.Normalize();

                    helper.RemoveDuplicatesFromColumn("B");

                    helper.DeleteColumn("A1");

                    helper.DeleteRow("A1");

                    helper.DeleteColumn("B1");

                    helper.SetColumnWidth(1, 18);

                    helper.SaveAsXLSX(Path.Combine(Environment.CurrentDirectory, $"{helper.GetLastRow()}.xlsx"));

                    File.Delete(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                    File.Delete(Path.Combine(Environment.CurrentDirectory, "messages.csv"));

                    stopWatch.Stop();

                    TimeSpan ts = stopWatch.Elapsed;

                    string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}", ts.Hours, ts.Minutes, ts.Seconds);

                    Console.Clear();

                    Console.WriteLine($"Обработка прошла успешно. Затраченное время: {elapsedTime}\nПроект: {projectName}\nКоличество чистых заявок: {helper.GetLastRow()}\n");

                    helper.Dispose();

                    Console.ReadLine();
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
