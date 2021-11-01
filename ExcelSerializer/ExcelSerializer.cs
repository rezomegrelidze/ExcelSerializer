using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSerializer
{
   public class ExcelSerializer : IDisposable
    {
        private Excel.Application excelType;

        public ExcelSerializer()
        {
            excelType = new Excel.Application { Visible = true };
        }

        public List<T> ExcelFileToData<T>(string filePath, int startRow = 0,int? endRow = null)
        {
            return WorksheetToData<T>(excelType.Workbooks.Open(filePath).Worksheets[1],startRow,endRow);
        }

        public Excel.Worksheet DisplayInExcel<T>(IEnumerable<T> data)
        {


            excelType.Workbooks.Add();
            var workSheet = (Excel.Worksheet)excelType.ActiveSheet;

            var type = typeof(T);
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var propertyNames = properties.Select(p => p.Name).ToArray();

            for (int i = 1; i <= propertyNames.Length; i++)
            {
                var name = propertyNames[i - 1];
                workSheet.Cells[1, i] = name;
            }

            int rowIndex = 2;
            foreach (var item in data)
            {
                var values = propertyNames.Select(name => type.GetProperty(name).GetValue(item)).ToArray();

                for (int i = 1; i <= values.Length; i++)
                {
                    workSheet.Cells[rowIndex, i] = values[i - 1];
                }

                rowIndex++;
            }

            var usedrange = workSheet.UsedRange;
            usedrange.Columns.AutoFit();
            usedrange.Rows.AutoFit();
            return workSheet;
        }

        public List<T> WorksheetToData<T>(Excel.Worksheet workSheet,int startRow = 0,int? endRow = null)
        {
            var data = new List<T>();
            var type = typeof(T);
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var propertyNames = properties.Select(p => p.Name).ToArray();

            Excel.Range userRange = workSheet.UsedRange;
            int columnCount = userRange.Columns.Count;
            int start = startRow;
            int end = endRow ?? userRange.Rows.Count;


            for (int row = start; row <= end; row++)
            {
                var values = new Queue<object>();
                for (int col = 1; col <= columnCount; col++)
                {
                    values.Enqueue(((Excel.Range)workSheet.Cells[row, col]).Value);
                }
                dynamic obj = Activator.CreateInstance<T>();
                foreach (var propertyName in propertyNames)
                {
                    dynamic value = values.Dequeue();
                    typeof(T).GetProperty(propertyName).SetValue(obj, value);
                }
                data.Add((T)obj);
            }

            return data;
        }

        private void ReleaseUnmanagedResources()
        {
            Marshal.ReleaseComObject(excelType);
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        ~ExcelSerializer()
        {
            ReleaseUnmanagedResources();
        }
    }
}
