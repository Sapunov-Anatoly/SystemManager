using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Management;
using System.IO;
using System.Diagnostics;

namespace SystemManager
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("[+] Получение версии системы...");
            string systemVersion = Environment.OSVersion.ToString();

            Console.WriteLine("[+] Получение имени устройства...");
            string systemMachineName = Environment.MachineName;

            Console.WriteLine("[+] Получение имени и типа процессора...");
            string systemProcessorName = GetSystemProcessorInfo("name");
            string systemProcessorType = GetSystemProcessorInfo("type");

            Console.WriteLine("[+] Получение объема оперативной памяти...");
            double systemRamSize = GetSystemRamSize();

            Console.WriteLine("[+] Получение разрешения экрана...");
            string systemScreenResolution = GetSystemScreenResolution();

            Console.WriteLine("[+] Генерация excel-файла...");
            string pathToFile = GenerateExcelFile(systemVersion, systemMachineName, systemProcessorName, systemProcessorType, systemRamSize, systemScreenResolution);

            if (pathToFile != null)
            {
                OpenFileQuestion(pathToFile);

                Console.WriteLine("[+] Программа успешно завершила работу");
            }
        }

        static string GetActualDateTime()
        {
            string dateTime = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss").Replace(":", "_");
            return dateTime;
        }

        static void OpenFileQuestion(string pathToFile)
        {
            Console.WriteLine("\nОткрыть файл? Y/n");

            string openFileChecker;

            while (true)
            {
                openFileChecker = Console.ReadLine();

                if (openFileChecker == "Y" || openFileChecker == "y")
                {
                    try
                    {
                        Console.WriteLine("[+] Открытие файла...");
                        Process.Start(pathToFile);
                        Console.WriteLine("[+] Файл успешно открыт");

                        break;
                    }
                    catch (Exception e)
                    {
                        Console.Write("[!] Открытие файла завершилось ошибкой: ");
                        Console.WriteLine(e.Message);

                        break;
                    }
                }
                else if (openFileChecker == "N" || openFileChecker == "n") { break; }
                else
                {
                    Console.WriteLine("[!] Введите корректное значение");
                    continue;
                }
            }
        }

        static string GenerateExcelFile(string systemVersion, string systemMachineName, string systemProcessorName, string systemProcessorType, double systemRamSize, string systemScreenResolution)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                // Создание листа
                ExcelWorksheet mainList = package.Workbook.Worksheets.Add("Характеристики ПК");

                mainList.Cells["A1"].Value = "Версия ОС";
                mainList.Cells["A2"].Value = "Имя устройства";
                mainList.Cells["A3"].Value = "Имя процессора";
                mainList.Cells["A4"].Value = "Тип процессора";
                mainList.Cells["A5"].Value = "RAM";
                mainList.Cells["A6"].Value = "Разрешение экрана";

                mainList.Cells["B1"].Value = systemVersion;
                mainList.Cells["B2"].Value = systemMachineName;
                mainList.Cells["B3"].Value = systemProcessorName;
                // Чтобы экселька не ругалась на тип данных, проще всего конвертировать тут
                mainList.Cells["B4"].Value = Convert.ToDouble(systemProcessorType);
                mainList.Cells["B5"].Value = systemRamSize;
                mainList.Cells["B6"].Value = systemScreenResolution;

                // Сохранение Excel-файла
                try
                {
                    string getPathToExe = Environment.CurrentDirectory.ToString();
                    string actualDateTime = GetActualDateTime();
                    string pathToFile = getPathToExe + "\\Excels\\[" + actualDateTime + "]_Характеристики.xlsx";

                    FileInfo file = new FileInfo(pathToFile);
                    package.SaveAs(file);

                    Console.WriteLine("[+] Генерация excel-файла выполнена успешно");
                    Console.WriteLine("[+] Excel-файл находится по пути: " + pathToFile);

                    return pathToFile;
                }
                catch(Exception e)
                {
                    Console.Write("[!] Генерация excel-файла завершилась ошибкой: ");
                    Console.WriteLine(e.Message);
                    Console.WriteLine("[!] Возможно, удалена папка 'Excels'");
                    return null;
                }
            }
        }

        static string GetSystemScreenResolution()
        {
            Screen systemScreen = Screen.PrimaryScreen;
            int screenWidth = systemScreen.Bounds.Width;
            int screenHeight = systemScreen.Bounds.Height;

            return (screenWidth.ToString() + " x " + screenHeight.ToString());
        }

        static double GetSystemRamSize()
        {
            double systemRamSize = 0;

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");

            foreach (ManagementObject obj in searcher.Get())
            {
                ulong totalMemoryBytes = (ulong)obj["TotalPhysicalMemory"];
                systemRamSize = Math.Round(totalMemoryBytes / (1024.0 * 1024.0 * 1024.0), 2);
            }

            return systemRamSize;
        }

        static string GetSystemProcessorInfo(string findedInfo)
        {
            string systemProcessorName = "Неизвестно";
            string systemProcessorType = "Неизвестно";

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");

            foreach (ManagementObject obj in searcher.Get())
            {
                systemProcessorName = obj["Name"].ToString();
                systemProcessorType = obj["ProcessorType"].ToString();
            }

            switch (findedInfo)
            {
                case "name":

                    return systemProcessorName;

                case "type":

                    return systemProcessorType;
            }

            return "Неизвестно";
        }
    }
}
