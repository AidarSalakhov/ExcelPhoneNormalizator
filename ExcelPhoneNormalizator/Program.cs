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
                    if (helper.Open(filePath: Path.Combine(Environment.CurrentDirectory, "ForNormalization.xlsx")))
                    {
                        helper.removeDuplicatesA();

                        Console.WriteLine("Нормализация телефонов...");
                        helper.Normalize();

                        helper.removeDuplicatesB();

                        helper.DeleteColumnA();

                        Console.WriteLine("Сохранение...");
                        helper.Save();

                        Console.WriteLine("Закрытие Excel...");
                        helper.Dispose();

                        Console.WriteLine("Готово!");
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
