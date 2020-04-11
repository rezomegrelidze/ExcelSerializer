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

            Excel.Range usedrange = workSheet.UsedRange;
            usedrange.Columns.AutoFit();
            usedrange.Rows.AutoFit();
            return workSheet;
        }

        public IEnumerable<T> WorksheetToData<T>(Excel.Worksheet workSheet)
        {
            var type = typeof(T);
            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            var propertyNames = properties.Select(p => p.Name).ToArray();

            Excel.Range userRange = workSheet.UsedRange;
            int columnCount = userRange.Columns.Count;
            int rowCount = userRange.Rows.Count;


            for (int row = 2; row <= rowCount; row++)
            {
                var values = new Queue<object>();
                for (int col = 1; col <= columnCount; col++)
                {
                    values.Enqueue((workSheet.Cells[row, col]
                        as Microsoft.Office.Interop.Excel.Range).Value);
                }
                dynamic obj = Activator.CreateInstance<T>();
                foreach (var propertyName in propertyNames)
                {
                    dynamic value = values.Dequeue();
                    typeof(T).GetProperty(propertyName).SetValue(obj, value);
                }
                yield return (T)obj;
            }
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
