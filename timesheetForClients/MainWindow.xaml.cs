﻿using System.IO;
using System.Text;
//using Syncfusion.XlsIO;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System;
using System.Linq;
//using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.Windows.Shapes;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Drawing;

namespace timesheetForClients
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string selectedFolder = "";
        string dateStart = "";
        XLColor colorHeaderB;
        XLColor colorDayB;
        XLColor colorResultHeaderB;
        XLColor colorResultB;

        public MainWindow()
        {
            InitializeComponent();
            //selectedFolder = "C:\\Users\\User\\Desktop\\тест";
            // dateStart = "30.12.2024 00:00:00";
        }
        private void buttonSelectCatalog_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() == true)
            {
                selectedFolder = dialog.FolderName;
                catalogNameLabel.Content = selectedFolder;
            }
        }

        private void buttonForm_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(selectedFolder) && !string.IsNullOrEmpty(dateStart))
            {
                var files = Directory.GetFiles(selectedFolder, "*employee*.xlsm").ToList();
                HashSet<string> uniqueProjects = new HashSet<string>();

                // Сбор уникальных проектов
                foreach (var file in files)
                {
                    using (var workbook = new XLWorkbook(file))
                    {
                        var worksheet = workbook.Worksheet(1); // Получаем первый лист
                        int rowCount = worksheet.LastRowUsed().RowNumber(); // Получаем количество строк
                        int startRow = -1;
                        colorHeaderB = XLColor.FromColor(worksheet.Cell("C10").Style.Fill.BackgroundColor.Color);
                        colorDayB = XLColor.FromColor(worksheet.Cell("C14").Style.Fill.BackgroundColor.Color);
                        colorResultHeaderB = XLColor.FromColor(worksheet.Cell("O10").Style.Fill.BackgroundColor.Color);
                        colorResultB = XLColor.FromColor(worksheet.Cell("O13").Style.Fill.BackgroundColor.Color);
                        //MessageBox.Show(colorHeaderB);
                        for (int row = 1; row <= rowCount; row++)
                        {
                            var cellValue = worksheet.Cell(row, 15).GetString(); // Получаем значение ячейки
                            if (!string.IsNullOrEmpty(cellValue) && cellValue.Trim() == dateStart.Trim())
                            {
                                startRow = row; // Запоминаем строку, где найдено совпадение
                                break; // Прерываем цикл, так как первое совпадение найдено
                            }
                        }

                        if (startRow != -1) // Если совпадение найдено, собираем названия проектов
                        {
                            for (int row = startRow + 1; row <= rowCount; row++) // Начинаем с строки под найденной
                            {
                                var projectNameCell = worksheet.Cell(row, 3).GetString().Trim(); // Получаем значение ячейки с названием проекта
                                if (!string.IsNullOrEmpty(projectNameCell) && !projectNameCell.Equals("ПРОЕКТ", StringComparison.OrdinalIgnoreCase))
                                    uniqueProjects.Add(projectNameCell); // Добавляем проект в коллекцию уникальных проектов
                            }
                        }
                    }
                }

                // Создание таймшитов 
                foreach (var projectName in uniqueProjects)
                {
                    string newFileName = System.IO.Path.Combine(selectedFolder, $"timesheet_customer_{projectName}.xlsx");
                    if (!File.Exists(newFileName))
                        createTimesheet(newFileName, projectName, dateStart);
                }

                // Теперь проходим по файлам с неделями и копируем строки с задачами в файлы проектов
                foreach (var file in files)
                {
                    using (var workbook = new XLWorkbook(file))
                    {
                        var worksheet = workbook.Worksheet(1);
                        int rowCount = worksheet.LastRowUsed().RowNumber();
                        for (int row = 1; row <= rowCount; row++)
                        {
                            var projectNameCell = worksheet.Cell(row, 3).GetString().Trim(); // Получаем значение ячейки с названием проекта
                            if (uniqueProjects.Contains(projectNameCell))
                                copyTaskRow(projectNameCell, worksheet, row);
                        }
                    }
                }
                foreach (var projectName in uniqueProjects)
                {
                    formingTheResults(projectName);
                }
            }
        }
        private void createTimesheet(string filePath, string projectName, string dateStart)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Timesheet");

                // Шапка
                worksheet.Cell(2, 4).Value = "ТАБЛИЦА УЧЁТА РАБОЧЕГО ВРЕМЕНИ";
                worksheet.Range("D2:G2").Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range("D2:G2").Merge().Style.Font.FontSize = 20;
                worksheet.Range("D2:G2").Style.Font.Bold = true;
                for (int i = 4; i <= 8; i++)
                    worksheet.Column(i).Width = 15;
                for (int i = 1; i <= 3; i++)
                    worksheet.Column(i).Width = 1;
                worksheet.Cell(4, 4).Value = "Проект";
                worksheet.Cell(4, 5).Value = projectName;
                worksheet.Cell(6, 13).Value = "ДАТА НАЧАЛА НЕДЕЛИ";
                worksheet.Range("M6:N6").Merge();
                worksheet.Cell(6, 13).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(6, 15).Value = dateStart.Split(' ')[0];
                worksheet.Cell(6, 15).Style.DateFormat.Format = "dd/MM/yyyy";
                worksheet.Cell(6, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Заголовок таблицы
                worksheet.Cell(8, 4).Value = "ЗАДАЧА";
                worksheet.Cell(8, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(8, 4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("D8:G9").Merge();
                worksheet.Cell(8, 4).Style.Fill.BackgroundColor = colorHeaderB;
                worksheet.Cell(8, 4).Style.Font.FontColor = XLColor.White;
                worksheet.Cell(8, 15).Value = "ИТОГО ЧАСОВ";
                worksheet.Cell(8, 15).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range("O8:O9").Merge();
                worksheet.Cell(8, 15).Style.Fill.BackgroundColor = colorResultHeaderB;
                worksheet.Cell(8, 15).Style.Font.FontColor = XLColor.White;
                worksheet.Column(15).Width = 15;
                worksheet.Cell(8, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // дни
                var daysOfWeek = new[] { "ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ" };

                for (int i = 0; i < daysOfWeek.Length; i++)
                {
                    // дни
                    worksheet.Cell(8, 8 + i).Value = daysOfWeek[i];
                    worksheet.Cell(8, 8 + i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(8, 8 + i).Style.Fill.BackgroundColor = colorHeaderB;
                    worksheet.Cell(8, 8 + i).Style.Font.FontColor = XLColor.White;
                    // даты
                    worksheet.Cell(9, 8 + i).Style.DateFormat.Format = "dd/MM/yyyy";
                    worksheet.Cell(9, 8 + i).FormulaA1 = $"=O6+{i}";
                    worksheet.Cell(9, 8 + i).Style.Fill.BackgroundColor = colorDayB;
                    worksheet.Cell(9, 8 + i).Style.Font.FontColor = XLColor.White;
                    worksheet.Cell(9, 8 + i).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                }
                worksheet.Column(8).AdjustToContents();
                workbook.SaveAs(filePath);
            }
        }
        private void copyTaskRow(string projectName, IXLWorksheet worksheet, int row)
        {
            string newFileName = System.IO.Path.Combine(selectedFolder, $"timesheet_customer_{projectName}.xlsx");
            if (File.Exists(newFileName))
            {
                using (var workbook = new XLWorkbook(newFileName))
                {
                    var worksheetProject = workbook.Worksheets.First(); // Получаем первый лист
                    int lastRow = worksheetProject.LastRowUsed().RowNumber(); // Получаем количество строк
                    var taskName = worksheet.Cell(row, 4).GetString().Trim(); // Получаем значение ячейки с названием задачи из текущей строки
                    bool taskExists = false;

                    for (int i = 1; i <= lastRow; i++)
                    {
                        var existingTaskName = worksheetProject.Cell(i, 4).GetString().Trim(); // Получаем значение ячейки с названием задачи
                        if (existingTaskName.Equals(taskName, StringComparison.OrdinalIgnoreCase))
                        {
                            taskExists = true;
                            // Суммируем часы
                            for (int j = 8; j <= 14; j++)
                            {
                                double existingHours = 0;
                                // Используем TryGetValue для безопасного получения существующих часов
                                if (worksheetProject.Cell(i, j).TryGetValue<double>(out existingHours))
                                {
                                    double newHours = 0;

                                    // Используем TryGetValue для безопасного получения нового значения из следующей строки
                                    if (worksheet.Cell(row + 1, j).TryGetValue<double>(out newHours)) // Изменено на row + 1
                                        worksheetProject.Cell(i, j).Value = existingHours + newHours; // Суммируем часы
                                }
                            }
                            break;
                        }
                    }

                    if (!taskExists)
                    {
                        // Копируем строку с задачей в файл проекта в 4-й столбец
                        worksheetProject.Cell(lastRow + 1, 4).Value = taskName; // Название задачи в 4-й столбец

                        worksheetProject.Range(worksheetProject.Cell(lastRow + 1, 4), worksheetProject.Cell(lastRow + 1, 7)).Merge().Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Копируем часы для каждого дня из следующей строки в столбцы H (8) по N (14)
                        for (int j = 8; j <= 14; j++)
                        {
                            worksheetProject.Column(j).Width = 15;
                            double hours = 0;
                            // Проверяем, является ли значение числом
                            if (worksheet.Cell(row + 1, j).TryGetValue<double>(out hours)) // Изменено на row + 1
                                worksheetProject.Cell(lastRow + 1, j).Value = hours; // Копируем часы для каждого дня из следующей строки
                            else
                                worksheetProject.Cell(lastRow + 1, j).Value = 0; // Если значение не число, устанавливаем 0
                            worksheetProject.Cell(lastRow + 1, j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }
                        worksheetProject.Cell(lastRow + 1, 15).FormulaA1 = $"SUM(H{lastRow + 1}:N{lastRow + 1})";
                        worksheetProject.Cell(lastRow + 1, 15).Style.Fill.BackgroundColor = colorResultB;
                        worksheetProject.Cell(lastRow + 1, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    }
                    workbook.Save();
                }
            }
        }
        private void formingTheResults(string projectName)
        {
            string newFileName = System.IO.Path.Combine(selectedFolder, $"timesheet_customer_{projectName}.xlsx");
            if (File.Exists(newFileName))
            {
                using (var workbook = new XLWorkbook(newFileName))
                {
                    var worksheetProject = workbook.Worksheets.First();
                    int lastRow = worksheetProject.LastRowUsed().RowNumber();
                    worksheetProject.Cell(lastRow + 1, 4).Value = "ИТОГО ЧАСОВ";
                    worksheetProject.Range($"D{lastRow + 1}:G{lastRow + 1}").Merge();
                    worksheetProject.Cell(lastRow + 1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheetProject.Cell(lastRow + 1, 4).Style.Fill.BackgroundColor = colorDayB;
                    worksheetProject.Cell(lastRow + 1, 4).Style.Font.FontColor = XLColor.White;
                    // Суммируем часы по столбцам с H (8) по N (14)
                    for (int j = 8; j <= 15; j++)
                    {
                        // Формула для суммирования значений от 10 строки до последней
                        string columnLetter = worksheetProject.Column(j).ColumnLetter();
                        worksheetProject.Cell(lastRow + 1, j).FormulaA1 = $"SUM({columnLetter}10:{columnLetter}{lastRow})";
                        worksheetProject.Cell(lastRow + 1, j).Style.Fill.BackgroundColor = colorDayB;
                        worksheetProject.Cell(lastRow + 1, j).Style.Font.FontColor = XLColor.White;
                        worksheetProject.Cell(lastRow + 1, j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    }
                    worksheetProject.Cell(lastRow + 1, 15).Style.Fill.BackgroundColor = colorResultHeaderB;

                    string columnLetterForResult = worksheetProject.Column(15).ColumnLetter();
                    int startRow = 8;
                    int endRow = lastRow + 1;

                    // Убедитесь, что columnLetterForResult не пустой и корректен
                    if (!string.IsNullOrEmpty(columnLetterForResult))
                    {
                        // Устанавливаем стиль для границ
                        var range = worksheetProject.Range($"D{startRow}:{columnLetterForResult}{endRow}");
                        range.Style.Border.RightBorderColor = XLColor.LightGray;
                        range.Style.Border.LeftBorderColor = XLColor.LightGray;
                        range.Style.Border.TopBorderColor = XLColor.LightGray;
                        range.Style.Border.BottomBorderColor = XLColor.LightGray;
                    }
                    foreach (var cell in worksheetProject.Cells())
                    {
                        cell.Style.Font.FontName = "Calibri";
                        cell.Style.Font.Bold = true;
                    }
                    workbook.Save();
                }
            }
        }
        private void StartDayWeekCalendar_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            dateStart = StartDayWeekCalendar.SelectedDate.Value.Date.ToString();
            //MessageBox.Show(dateStart);
        }
    }
}