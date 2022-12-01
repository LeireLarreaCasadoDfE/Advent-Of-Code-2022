using Microsoft.Office.Interop.Excel;
using System.Collections.Immutable;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using Range = Microsoft.Office.Interop.Excel.Range;

class Program
{
    static void Main(string[] args)
    {
        double currentCalories = 0.00;
        double maxCalories = 0.00;

        SetAllExcelClosed();
        
        //Create COM Objects.
        Application excelApp = new Application();
        if (excelApp == null)
        {
            Console.WriteLine("Excel is not installed!!");
            return;
        }
        Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\LLARREACASADO\OneDrive - Department for Education\Downloads\adventofcode01-01data2.xlsx");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Range excelRange = excelSheet.UsedRange;

        double topCaloriePack = GetTopCaloriePack(currentCalories, maxCalories, excelRange);
        Console.WriteLine($"This Elf is the top Keto Killer: {topCaloriePack.ToString()}");

        double top3CaloriesPack = GetTop3CaloriesPack(currentCalories, maxCalories, excelRange);
        Console.WriteLine($"And the ones getting ready for winter total a kcal count of: {top3CaloriesPack.ToString()}");

        SetExcelAppClean(excelBook, excelApp);
    }

    private static double GetTopCaloriePack(double currentCalories, double maxCalories, Range excelRange)
    {
        for (int i = 1; i <= excelRange.Rows.Count; i++)
        {
            if (excelRange.Cells[i].Value != null)
            {
                currentCalories = currentCalories + excelRange.Cells[i].Value;
            }
            else
            {
                if (currentCalories > maxCalories)
                {
                    maxCalories = currentCalories;
                }
                currentCalories = 0.00;
            }
        }
        return maxCalories;
    }

    private static double GetTop3CaloriesPack(double currentCalories, double maxCalories, Range excelRange)
    {
        //Get 3 MaxCalories
        List<double> listOfCalories = new List<double>();

        for (int i = 1; i <= excelRange.Rows.Count; i++)
        {
            if (excelRange.Cells[i].Value != null)
            {
                currentCalories = currentCalories + excelRange.Cells[i].Value;
            }
            else
            {
                listOfCalories.Add(currentCalories);
                currentCalories = 0.00;
            }
        }

        listOfCalories.Sort();
        maxCalories = listOfCalories[listOfCalories.Count - 1] + listOfCalories[listOfCalories.Count - 2] + listOfCalories[listOfCalories.Count - 3];

        return maxCalories;
    }

    private static void SetAllExcelClosed()
    {
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            process.Kill();
        }
    }

    private static void SetExcelAppClean(Workbook excelBook, Application excelApp)
    {
        excelBook.Close();
        excelApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
    }
}
