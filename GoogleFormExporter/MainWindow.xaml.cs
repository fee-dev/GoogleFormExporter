using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace GoogleFormExporter
{
    public partial class MainWindow 
    {
        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Google Form\\");
            var workBook = new ExcelPackage(new FileInfo(Location.Text));
            var workSheet = workBook.Workbook.Worksheets.First();
            var rowEnd = workSheet.Dimension.End.Row;
            var columnEnd = workSheet.Dimension.End.Column;
            var list = new List<Person>();
            for (var column = 4; column <= columnEnd; column++)
            {
                if (workSheet.Cells[1, column].Value != null && workSheet.Cells[1, column].Value.ToString().Contains("/"))
                {
                    for (int row = 2; row <= rowEnd; row++)
                    {
                        var riceValue = workSheet.Cells[row, column + 1].Value.ToString();
                        var resultString = Regex.Match(riceValue, @"\d+").Value;
                        var rice = !string.IsNullOrWhiteSpace(resultString) ? int.Parse(resultString) : 0;
                        list.Add(new Person
                        {
                            RoomNumber = int.Parse(workSheet.Cells[row, 4].Value.ToString()),
                            Name = workSheet.Cells[row, 2].Value.ToString(),
                            Date = workSheet.Cells[1, column].Value.ToString(),
                            Food = workSheet.Cells[row, column].Value.ToString(),
                            Rice = rice
                        });
                    }
                }
            }
            using (var excel = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excel.Workbook.Properties.Author = "Bryan";
                excel.Workbook.Properties.Title = "Google Form Data";
                excel.Workbook.Properties.Subject = "Google Form Data";
                excel.Workbook.Properties.Created = DateTime.Now;
                list = list.OrderBy(x=>x.Date).ThenBy(x => x.RoomNumber).ToList();
                var dates = list.Select(x => x.Date).OrderBy(x=>x).Distinct().ToList();
                foreach (var date in dates)
                {
                    //var foods = list.Where(x=>x.Date == date && x.Food != "NO ORDER FOR THIS DAY" && !x.Food.Contains(",")).Select(x => x.Food).Distinct().ToList();
                    var foodss = list.Where(x=>x.Date == date && x.Food != "NO ORDER FOR THIS DAY").Select(x => x.Food.Split(',').Where(y=>y!= "NO ORDER FOR THIS DAY")).ToList();
                    var foods = new List<string>();
                    foreach (var items in foodss)
                    {
                        foreach (var item in items)
                        {
                            foods.Add(item.Trim());
                        }
                    }
                    var rooms = list.Where(x => x.Date == date).Select(x => x.RoomNumber).Distinct().ToList();
                    var sheet = excel.Workbook.Worksheets.Add(date);
                    sheet.Cells[1, 1].Value = date;
                    var row = 1;
                    foreach (var food in foods.Select(x => x.Substring(0, x.IndexOf(x.ToCharArray().First(char.IsDigit)))).Distinct())
                    {
                        row++;
                    }
                    sheet.Cells[5, 4].Formula = "=SUM(D1:D4)";
                    sheet.Cells[5, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[5, 4].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                    sheet.Cells[6, 3].Value = "RICE";
                    sheet.Cells[6, 4].Value = list.Where(x => x.Date == date).Sum(x=>x.Rice);
                    sheet.Cells[8, 4].Value = "Ulam Quantity";
                    sheet.Cells[8, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[8, 4].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    sheet.Cells[8, 6].Value = "Rice Quantity";
                    sheet.Cells[8, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[8, 6].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    sheet.Cells["A1:F6"].Style.Font.Bold = true;
                    row = 9;
                    var roomStart = 9;
                    Person previousValue = null;
                    foreach (var room in rooms)
                    {
                        foreach (var person in list.Where(x=>x.Date == date && x.RoomNumber == room))
                        {
                            var currentRow = row - 1;
                            var rowFinal = row;
                            var rowsToAdd = 1;
                            if (row % 56 == 0)
                            {
                                while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                {
                                    rowsToAdd++;
                                    currentRow--;
                                }
                                sheet.InsertRow(currentRow + 1, rowsToAdd);
                                row = rowFinal + rowsToAdd;
                            }
                            else if (row % 57 == 0)
                            {
                                while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                {
                                    rowsToAdd++;
                                    currentRow--;
                                }
                                sheet.InsertRow(currentRow + 1, rowsToAdd);
                                row = rowFinal + rowsToAdd;
                            }
                            if (previousValue == null || person.RoomNumber == previousValue.RoomNumber)
                            {
                                if (person.Food.Contains(","))
                                {
                                    var foodList = person.Food.Split(',');
                                    sheet.Cells[row, 1].Value = person.RoomNumber;
                                    sheet.Cells[row, 2].Value = person.Name;
                                    sheet.Cells[row, 6].Value = person.Rice;
                                    row = row - 1;
                                    foreach (var food in foodList.Where(x=>!x.Equals("NO ORDER FOR THIS DAY")))
                                    {
                                        currentRow = row - 1;
                                        rowFinal = row;
                                        rowsToAdd = 1;
                                        if (row % 55 == 0)
                                        {
                                            while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                            {
                                                rowsToAdd++;
                                                currentRow--;
                                            }
                                            sheet.InsertRow(currentRow + 1, rowsToAdd);
                                            row = rowFinal + rowsToAdd;
                                        }
                                        if (row % 56 == 0)
                                        {
                                            while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                            {
                                                rowsToAdd++;
                                                currentRow--;
                                            }
                                            sheet.InsertRow(currentRow + 1, rowsToAdd);
                                            row = rowFinal + rowsToAdd;
                                        }
                                        else if (row % 57 == 0)
                                        {
                                            while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                            {
                                                rowsToAdd++;
                                                currentRow--;
                                            }
                                            sheet.InsertRow(currentRow + 1, rowsToAdd);
                                            row = rowFinal + rowsToAdd;
                                        }
                                        row++;
                                        sheet.Cells[row, 3].Value = food.Substring(0, food.IndexOf(food.ToCharArray().First(char.IsDigit))).Trim();
                                        sheet.Cells[row, 4].Value = int.Parse(food.Substring(food.IndexOf(food.ToCharArray().First(char.IsDigit)), 1));
                                        //sheet.Cells[row, 6].Value = 0;
                                    }
                                }
                                else
                                {
                                    sheet.Cells[row, 1].Value = person.RoomNumber;
                                    sheet.Cells[row, 2].Value = person.Name;
                                    var a = person.Food.Equals("NO ORDER FOR THIS DAY") ? "" :person.Food.Substring(0, person.Food.IndexOf(person.Food.ToCharArray().First(char.IsDigit))).Trim();
                                    sheet.Cells[row, 3].Value =  a;
                                    if (person.Food.Equals("NO ORDER FOR THIS DAY"))
                                        sheet.Cells[row, 4].Value = "";
                                    else
                                        sheet.Cells[row, 4].Value = int.Parse(person.Food.Substring(person.Food.IndexOf(person.Food.ToCharArray().First(char.IsDigit)), 1));
                                    sheet.Cells[row, 6].Value = person.Rice;
                                }
                            }
                            else
                            {
                                sheet.Cells[row, 4].Formula = "=SUM(D" + roomStart + ":D" + (row - 1) + ")";
                                sheet.Cells[row, 4].Style.Font.Color.SetColor(Color.White);
                                sheet.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.Black);
                                sheet.Cells[row, 6].Formula = "=SUM(F" + roomStart + ":F" + (row - 1) + ")";
                                sheet.Cells[row, 6].Style.Font.Color.SetColor(Color.White);
                                sheet.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                sheet.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(Color.Black);
                                row++;
                                row++;
                                roomStart = row;
                                if (person.Food.Contains(","))
                                {
                                    var foodList = person.Food.Split(',');
                                    sheet.Cells[row, 1].Value = person.RoomNumber;
                                    sheet.Cells[row, 2].Value = person.Name;
                                    sheet.Cells[row, 6].Value = person.Rice;
                                    row = row - 1;
                                    foreach (var food in foodList.Where(x => !x.Equals("NO ORDER FOR THIS DAY")))
                                    {
                                        currentRow = row - 1;
                                        rowFinal = row;
                                        rowsToAdd = 1;
                                        if (row % 55 == 0)
                                        {
                                            while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                            {
                                                rowsToAdd++;
                                                currentRow--;
                                            }
                                            sheet.InsertRow(currentRow + 1, rowsToAdd);
                                            row = rowFinal + rowsToAdd;
                                        }
                                        if (row % 56 == 0)
                                        {
                                            while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                            {
                                                rowsToAdd++;
                                                currentRow--;
                                            }
                                            sheet.InsertRow(currentRow + 1, rowsToAdd);
                                            row = rowFinal + rowsToAdd;
                                        }
                                        else if (row % 57 == 0)
                                        {
                                            while (!string.IsNullOrWhiteSpace(sheet.Cells[currentRow, 3].Value?.ToString()))
                                            {
                                                rowsToAdd++;
                                                currentRow--;
                                            }
                                            sheet.InsertRow(currentRow + 1, rowsToAdd);
                                            row = rowFinal + rowsToAdd;
                                        }
                                        row++;
                                        sheet.Cells[row, 3].Value = food.Substring(0, food.IndexOf(food.ToCharArray().First(char.IsDigit))).Trim();
                                        sheet.Cells[row, 4].Value = int.Parse(food.Substring(food.IndexOf(food.ToCharArray().First(char.IsDigit)), 1));
                                        //sheet.Cells[row, 6].Value = 0;
                                    }
                                }
                                else
                                {
                                    sheet.Cells[row, 1].Value = person.RoomNumber;
                                    sheet.Cells[row, 2].Value = person.Name;
                                    sheet.Cells[row, 3].Value = person.Food.Equals("NO ORDER FOR THIS DAY") ? "" : person.Food.Substring(0, person.Food.IndexOf(person.Food.ToCharArray().First(char.IsDigit))).Trim();
                                    if (person.Food.Equals("NO ORDER FOR THIS DAY"))
                                        sheet.Cells[row, 4].Value = "";
                                    else
                                        sheet.Cells[row, 4].Value = int.Parse(person.Food.Substring(person.Food.IndexOf(person.Food.ToCharArray().First(char.IsDigit)), 1));
                                    sheet.Cells[row, 6].Value = person.Rice;
                                }
                            }

                            previousValue = person;
                            row++;
                        }
                    }

                    sheet.Cells[row, 4].Formula = "=SUM(D" + roomStart + ":D" + (row - 1) + ")";
                    sheet.Cells[row, 4].Style.Font.Color.SetColor(Color.White);
                    sheet.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(Color.Black);
                    sheet.Cells[row, 6].Formula = "=SUM(F" + roomStart + ":F" + (row - 1) + ")";
                    sheet.Cells[row, 6].Style.Font.Color.SetColor(Color.White);
                    sheet.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(Color.Black);
                    var lastRow = row - 1;
                    row = 1;
                    foreach (var food in foods.Select(x => x.Substring(0, x.IndexOf(x.ToCharArray().First(char.IsDigit)))).Distinct())
                    {
                        var current = food;
                        var sum = foods.Where(x => x.Contains(food)).Select(x => int.Parse(x.Substring(x.IndexOf(x.ToCharArray().First(char.IsDigit)), 1))).Sum();
                        sheet.Cells[row, 3].Value = current;
                        sheet.Cells[row, 4].Value = sum;
                        row++;
                    }

                    lastRow++;
                    sheet.Cells["A1:F" + lastRow].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells["A1:F" + lastRow].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells["A1:F" + lastRow].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells["A1:F" + lastRow].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.PrinterSettings.TopMargin = decimal.Parse("0.17");
                    sheet.PrinterSettings.BottomMargin = decimal.Parse("0.17");
                    sheet.PrinterSettings.LeftMargin = decimal.Parse("0.24");
                    sheet.PrinterSettings.RightMargin = decimal.Parse("0.24");
                    sheet.Column(1).AutoFit(0);
                    sheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Column(2).AutoFit(0);
                    sheet.Column(3).AutoFit(0);
                    sheet.Column(3).Width = 40;
                    sheet.Column(4).AutoFit(0);
                    sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Column(5).Width = 2;
                    sheet.Column(6).AutoFit(0);
                    sheet.Column(6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }

                var excelFile = new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Google Form\\Google Form Data " + DateTime.Now.ToString("MM-dd-yyyy") + ".xlsx");
                excel.SaveAs(excelFile);
                Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Google Form\\Google Form Data " + DateTime.Now.ToString("MM-dd-yyyy") + ".xlsx");
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            var x = openFileDialog.ShowDialog();
            if (x.HasValue) Location.Text = openFileDialog.FileName;
        }
    }

    public class Person
    {
        public int RoomNumber { get; set; }
        public string Name { get; set; }
        public string Food { get; set; }
        public int Rice { get; set; }
        public string Date { get; set; }
    }
}
