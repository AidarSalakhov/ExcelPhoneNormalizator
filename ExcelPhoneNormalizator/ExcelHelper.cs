﻿using System;
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

        internal void SaveAs()
        {
            _workbook.SaveAs(Path.Combine(Environment.CurrentDirectory, "leads.xlsx"), Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _filePath = null;
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
            int lastRow = _excel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            for (int i = 1; i < lastRow; i++)
            {
                var val = Get(column: "A", row: i);

                try
                {
                    string stringVal = Convert.ToString(val);

                    var value = string.Join("", stringVal.Where(c => char.IsDigit(c)));

                    StringBuilder charVal = new StringBuilder(value);

                    if (charVal.Length != 11)
                    {
                        continue;
                    }
                    else if (charVal[0] == '7' && charVal[1] == '9')
                    {
                        Set(column: "B", row: i, data: charVal.ToString());
                    }
                    else if (charVal[0] == '8' && charVal[1] == '9')
                    {
                        charVal[0] = '7';
                        Set(column: "B", row: i, data: charVal.ToString());
                    }

                    Console.WriteLine($"Удачно преобразованая строка {i}");
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }

            }

        }

        public void removeDuplicatesB()
        {

            Excel.Range range = _excel.Range[$"B1:B{_excel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row}", Type.Missing];

            range.RemoveDuplicates(_excel.Evaluate(1),
                Excel.XlYesNoGuess.xlNo);
        }

        public void removeDuplicatesA()
        {

            Excel.Range range = _excel.Range[$"A1:A{_excel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row}", Type.Missing];

            range.RemoveDuplicates(_excel.Evaluate(1),
                Excel.XlYesNoGuess.xlNo);
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

        public int LastRow()
        {
            int lastRow = _excel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            return lastRow;
        }

    }



}
