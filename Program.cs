//Подключение библиотек
using System.Management; //Библиотека для получения подробных данных о вашем ПК из самой Windows.
using System.Windows.Forms; // Для собирания информации об мониторе
using OfficeOpenXml; //Библиотека EPPlus для создания и взаимодействия с Excel
using OfficeOpenXml.Style;
using System.Drawing;
class Program
{
    [STAThread] // Необходимо для работы с Windows Forms
    static void Main()
    {
        ExcelPackage.License.SetNonCommercialPersonal("UserName"); // Без этого EPPlus не позволит собой пользоваться и даст ошибку
        var systemInfo = GetSystemInfo();
        string fileName = "system_info.xlsx";
        string folderPath = AppDomain.CurrentDomain.BaseDirectory;
        string fullPath = Path.Combine(folderPath, fileName);
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Характеристики ПК");
            worksheet.Cells[1, 1].Value = "Параметр";
            worksheet.Cells[1, 2].Value = "Значение";
            using (var range = worksheet.Cells[1, 1, 1, 2]) // Стили для ячеек
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(191, 169, 235));
                range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                range.Style.Font.Bold = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            int row = 2;
            foreach (var item in systemInfo)
            {
                worksheet.Cells[row, 1].Value = item.Key;
                worksheet.Cells[row, 2].Value = item.Value;
                using (var range = worksheet.Cells[row, 1, row, 2]) // Стиль для ячеек
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(135, 255, 175));
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                row++;
            }
            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            try
            {
                package.SaveAs(new FileInfo(fullPath));
                Console.WriteLine("Файл с системной информацией создан по пути:");
                Console.WriteLine(fullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Отказано в доступе к файлу. Возможно, файл открыт в другой программе или нет прав на запись.");
                Console.WriteLine("Подробности: " + ex.Message);
            }
        }
        Console.WriteLine("Нажмите любую клавишу, чтобы выйти...");
        Console.ReadKey();
    }
    static Dictionary<string, string> GetSystemInfo()
    {
        var info = new Dictionary<string, string>();
        info["Версия ОС"] = Environment.OSVersion.VersionString;
        info["Имя компьютера"] = Environment.MachineName;
        info["Процессор"] = GetProcessorName();
        info["Объем ОЗУ (GB)"] = GetTotalRAM();
        info["Разрешение экрана"] = GetScreenResolution();
        return info;
    }
    static string GetProcessorName()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("select Name from Win32_Processor");
            foreach (var item in searcher.Get())
                return item["Name"]?.ToString() ?? "Не найдено";
        }
        catch { }
        return "Не найдено";
    }
    static string GetTotalRAM()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("select TotalPhysicalMemory from Win32_ComputerSystem");
            foreach (var item in searcher.Get())
            {
                ulong memBytes = (ulong)item["TotalPhysicalMemory"];
                return (memBytes / 1024 / 1024 / 1024).ToString(); // Округляет до целого числа ГБ
            }
        }
        catch { }
        return "Не известно";
    }
    static string GetScreenResolution()
    {
        try
        {
            int width = Screen.PrimaryScreen.Bounds.Width;
            int height = Screen.PrimaryScreen.Bounds.Height;
            return $"{width}x{height}";
        }
        catch
        {
            return "Неизвестно";
        }
    }
}