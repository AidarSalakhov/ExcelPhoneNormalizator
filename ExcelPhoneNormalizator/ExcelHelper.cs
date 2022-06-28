using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelPhoneNormalizator
{
    class ExcelHelper : IDisposable
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private string _filePath;

        public ExcelHelper()
        {
            _excel = new Excel.Application();
        }

        internal bool Open(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    _workbook = _excel.Workbooks.Open(filePath);
                }
                else
                {
                    _workbook = _excel.Workbooks.Add();
                    _filePath = filePath;
                }

                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

       
        internal void Save()
        {
            if (!string.IsNullOrEmpty(_filePath))
            {
                _workbook.SaveAs(_filePath);
                _filePath = null;
            }
            else
            {
                _workbook.Save();
            }
        }

        internal bool Set(string column, int row, object data)
        {
            try
            {
                ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column] = data;
                return true;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return false;
        }

        internal object Get(string column, int row)
        {
            try
            {
                return ((Excel.Worksheet)_excel.ActiveSheet).Cells[row, column].Value2;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
            return null;
        }

        public void Dispose()
        {
            try
            {
                _workbook.Close();
                _excel.Quit();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        public void Normalize()
        {

            for (int i = 1; i < 1000; i++)
            {
                var val = Get(column: "A", row: i);

                string stringVal = Convert.ToString(val);

                var value = string.Join("", stringVal.Where(c => char.IsDigit(c)));

                Set(column: "B", row: i, data: value);
            }
        }

        public void From8to7()
        {
            for (int i = 1; i < 1000; i++)
            {
                var val = Get(column: "B", row: i);

                StringBuilder stringVal = new StringBuilder(Convert.ToString(val));

                if (stringVal[0] == 7 && stringVal[1] == 9)
                {
                    Set(column: "C", row: i, data: stringVal.ToString());
                }
                else if (stringVal[0] == 8 && stringVal[1] == 9)
                {
                    stringVal[0] = '7';
                    Set(column: "C", row: i, data: stringVal.ToString());
                }
                else
                {
                    Set(column: "C", row: i, data: "");
                }
                
            };
        }
    }

}
