using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Linq;

namespace TestOffice
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excel = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet worksheet = null;
            Excel.Range range = null;
            Excel.ChartObjects chartObjs = null;
            Excel.ChartObject chartObj = null;
            Excel.Chart chart = null;

            try
            {
                excel = new Excel.Application();
                excel.Visible = true;
                workbooks = excel.Workbooks;
                workbook = workbooks.Add();
                sheets = workbook.Sheets;
                worksheet = (Excel.Worksheet)sheets[1];
                worksheet.Cells[1, 1] = "Hello";

                var processes = Process.GetProcesses().OrderByDescending(p => p.WorkingSet64).Take(10);
                int i = 2;
                foreach (var p in processes)
                {
                    worksheet.Cells[i, 1] = p.ProcessName;
                    worksheet.Cells[i, 2] = p.WorkingSet64;
                    i++;
                }

                range = worksheet.Range["A1", "B" + (i - 1)];
                chartObjs = (Excel.ChartObjects)worksheet.ChartObjects(Type.Missing);
                chartObj = chartObjs.Add(60, 10, 300, 250);
                chart = chartObj.Chart;
                chart.SetSourceData(range);
                chart.ChartWizard(Source: range, Title: "Memory Usage in " + Environment.MachineName);
                chart.ChartStyle = 45;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                if (chart != null) Marshal.ReleaseComObject(chart);
                if (chartObj != null) Marshal.ReleaseComObject(chartObj);
                if (chartObjs != null) Marshal.ReleaseComObject(chartObjs);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                if (excel != null) Marshal.ReleaseComObject(excel);
            }
        }
    }
}
