using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.Common;
using Microsoft.Office.Interop.Excel;

namespace ExcelPhoneNormalizator
{
    class ExcelOperations : IDisposable
    {
        private Excel.Application _excel;
        private Excel.Workbook _workbook;
        private string _filePath;

        public ExcelOperations()
        {
            _excel = new Excel.Application();
        }

        internal bool OpenCSV(string filePath)
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
            _workbook.SaveAs(outputFile, Excel.XlFileFormat.xlUnicodeText, AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);
        }

        internal void SaveAsXLSX(string outputFile)
        {
            _workbook.SaveAs(outputFile, Excel.XlFileFormat.xlOpenXMLWorkbook, AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);
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
                _workbook.Close(true);
                _excel.Quit();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        public void Normalize()
        {
            int lastRow = GetLastRow();

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
                else if ((charVal[0] == '7' || charVal[0] == '8') && charVal[1] == '9')
                {
                    charVal[0] = '7';

                    string charValToString = charVal.ToString();

                    if (!IsTooManyRepeatingNumbers(charValToString))
                    {
                        Set(column: "B", row: i, data: charValToString);
                    }
                }

                double progress = Math.Round((Convert.ToDouble(i) / Convert.ToDouble(lastRow)) * 100);

                if (progress % 5 == 0)
                {
                    Console.Clear();
                    Console.WriteLine($"Обработка телефонов: {progress}%");
                }
            }

        }

        public void RemoveDuplicatesFromColumn(string column)
        {
            Excel.Range range = _excel.Range[$"{column}1:{column}{GetLastRow()}", Type.Missing];
            range.RemoveDuplicates(_excel.Evaluate(1), Excel.XlYesNoGuess.xlNo);
        }

        public void RemoveDuplicatesFromManyColumns(string firstColumn, string lastColumn)
        {
            Excel.Range range = _excel.Range[$"{firstColumn}1:{lastColumn}{GetLastRow()}", Type.Missing];
            range.RemoveDuplicates(_excel.Evaluate(1), Excel.XlYesNoGuess.xlNo);
        }

        public void DeleteColumn(string column)
        {
            Excel.Range range = _excel.get_Range(column, Type.Missing);
            range.EntireColumn.Delete(Type.Missing);
        }

        public void DeleteRow(string column)
        {
            Excel.Range range = _excel.get_Range(column, Type.Missing);
            range.EntireRow.Delete(Type.Missing);
        }

        public int GetLastRow()
        {
            int lastRealRow = _excel.Cells.Find("*", SearchOrder: Excel.XlSearchOrder.xlByRows, SearchDirection: Excel.XlSearchDirection.xlPrevious, MatchCase: false).Row;
            return lastRealRow;
        }

        public void SetColumnWidth(int column, int width)
        {
            _excel.ActiveSheet.Columns[column].ColumnWidth = width;
        }

        private bool IsTooManyRepeatingNumbers(string value)
        {
            if (value[4] == value[5] && value[5] == value[6] && value[6] == value[7])
            {
                return true;
            }
            return false;
        }

        public string GetProjectName()
        {
            string projectName = Convert.ToString(Get("A", 2));
            return projectName;
        }

        public bool HasOneKavichkuAsLastChar(string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                char lastChar = input[input.Length - 1];
                if (lastChar == '"')
                {
                    return input.Count(x => x == '"') == 1;
                }
            }
            return false;
        }

        public void Transpose()
        {
            int startValue = 1;

            int resultValue = 1;

            int lastRow = GetLastRow();

            for (int i = 1; i <= lastRow; i++)
            {
                var val = Get(column: "A", row: i);

                string stringVal = Convert.ToString(val);

                if (HasOneKavichkuAsLastChar(stringVal))
                {
                    Excel.Range rangeColumn = _excel.Range[$"A{startValue}:A{i}", Type.Missing];
                    rangeColumn.Copy();
                    Excel.Range rangeRow = _excel.Range[$"B{resultValue}", Type.Missing];
                    rangeRow.PasteSpecial(XlPasteType.xlPasteAll,XlPasteSpecialOperation.xlPasteSpecialOperationNone,true,true);
                    startValue = i+1;
                    resultValue++;
                }

                double progress = Math.Round((Convert.ToDouble(i) / Convert.ToDouble(lastRow)) * 100);

                if (progress % 5 == 0)
                {
                    Console.Clear();
                    Console.WriteLine($"Обработка файла: {progress}%");
                }
            }
        }

        public void DeleteBlankCells(string column)
        {
            Excel.Range range = _excel.Range[column, Type.Missing];
            Excel.Range blankCells = range.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
            Excel.Range rowsToDelete = blankCells.EntireRow;
            rowsToDelete.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
        }

        public void GoToTopLeft()
        {
            Excel.Range topLeft = _excel.Range["A1", Type.Missing];
            topLeft.Select();
            _excel.Goto(topLeft, Type.Missing);
        }
    }
}
