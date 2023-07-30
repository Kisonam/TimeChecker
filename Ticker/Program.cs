using OfficeOpenXml;
using System.Diagnostics;

internal class Program
{
    private static void Main(string[] args)
    {
        string appPath = "C:Your path"; // Замініть на шлях до вашої програми .exe
        string excelFilePath = "D:\\ExcelFile.xlsx"; // Замініть на шлях до ексель таблиці

        //// Створюємо ексель файл з необхідною структурою таблиці, якщо його ще не існує
        //if (!File.Exists(excelFilePath))
        //{
        //    using (ExcelPackage excelPackage = new ExcelPackage())
        //    {
        //        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("AppUsage");
        //        worksheet.Cells[1, 1].Value = "Days/Months";

        //        DateTime _currentDate = DateTime.Now;
        //        for (int i = 1; i <= 31; i++)
        //        {
        //            worksheet.Cells[i + 1, 1].Value = i;
        //        }

        //        for (int i = 0; i < 12; i++)
        //        {
        //            worksheet.Cells[1, i + 2].Value = _currentDate.AddMonths(i).ToString("MMMM");
        //        }

        //        excelPackage.SaveAs(new FileInfo(excelFilePath));
        //    }
        //}

        Process appProcess = null;
        DateTime startTime = DateTime.Now;
        DateTime currentDate = DateTime.Now;

        while (true)
        {
            if (appProcess == null || appProcess.HasExited)
            {
                // Запускаємо новий процес, якщо попередній закрився або програма ще не відкрита
                appProcess = Process.Start(appPath);
                startTime = DateTime.Now;
            }

            // Оновлюємо поточну дату
            currentDate = DateTime.Now;

            // Очікуємо 1 хвилину перед наступною перевіркою
            Thread.Sleep(60000);

            if (!appProcess.HasExited)
            {
                TimeSpan elapsedTime = currentDate - startTime;
                int row = currentDate.Day + 1;
                int col = (currentDate.Month - DateTime.Now.Month + 12) % 12 + 2;

                // Записуємо час роботи програми у відповідну комірку таблиці
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["AppUsage"];

                    // Перевіряємо, чи вже було записано число для поточного дня, якщо так - додаємо до нього нове значення
                    if (worksheet.Cells[row, col].Value != null)
                    {
                        double existingValue = (double)worksheet.Cells[row, col].Value;
                        worksheet.Cells[row, col].Value = existingValue + elapsedTime.TotalMinutes;
                    }
                    else
                    {
                        worksheet.Cells[row, col].Value = elapsedTime.TotalMinutes;
                    }

                    excelPackage.Save();
                }
            }
        }
    }
}