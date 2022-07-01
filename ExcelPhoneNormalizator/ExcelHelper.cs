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
                    _workbook = _excel.Workbooks.Open(filePath, Format: 6, Delimiter: ";");
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

        internal void SaveAsTXT(string outputFile)
        {
            _workbook.SaveAs(Filename: outputFile, FileFormat: Excel.XlFileFormat.xlUnicodeText, AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);
        }

        internal void SaveAsXLSX(string outputFile)
        {
            _workbook.SaveAs(outputFile, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
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
            int lastRow = LastRealRow();

            for (int i = 1; i < lastRow; i++)
            {
                var val = Get(column: "A", row: i);

                string stringVal = Convert.ToString(val);

                var value = string.Join("", stringVal.Where(c => char.IsDigit(c)));

                StringBuilder charVal = new StringBuilder(value);

                if (charVal.Length != 11)
                {
                    continue;
                }
                else if (charVal[0] == '7' && charVal[1] == '9' || charVal[0] == '8' && charVal[1] == '9')
                {
                    charVal[0] = '7';

                    string charValToString = charVal.ToString();

                    if (!IsTooManyRepeatingNumbers(charValToString))
                    {
                        Set(column: "B", row: i, data: charValToString);
                    }
                }

                Console.WriteLine($"Удачно преобразованая строка {i} из {lastRow}");
            }

        }

        public void RemoveDuplicatesB()
        {
            Excel.Range range = _excel.Range[$"B1:B{_excel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row}", Type.Missing];
            range.RemoveDuplicates(_excel.Evaluate(1), Excel.XlYesNoGuess.xlNo);
        }

        public void RemoveDuplicatesA()
        {
            Excel.Range range = _excel.Range[$"A1:A{_excel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row}", Type.Missing];
            range.RemoveDuplicates(_excel.Evaluate(1), Excel.XlYesNoGuess.xlNo);
        }

        public void DeleteColumn(string column)
        {
            Excel.Range range = _excel.get_Range(column, Type.Missing);
            range.EntireColumn.Delete(Type.Missing);
        }

        public void DeleteEntireRow(string column)
        {
            Excel.Range range = _excel.get_Range(column, Type.Missing);
            range.EntireRow.Delete(Type.Missing);
        }

        public int LastRealRow()
        {
            int lastRealRow = _excel.Cells.Find("*", Type.Missing, Type.Missing, Type.Missing, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing).Row;
            return lastRealRow;
        }

        public void SetColumnWidth(int column, int width)
        {
            _excel.ActiveSheet.Columns[column].ColumnWidth = width;
        }

        public bool IsTooManyRepeatingNumbers(string value)
        {
            if (value[4] == value[5] && value[5] == value[6] && value[6] == value[7])
            {
                return true;
            }

            return false;
        }

        public void Speed()
        {
            //speed = speedNow * 0.5 + speedLastHalfMinute * 0.3 + speedLastMinute * 0.2
        }


    }
}
