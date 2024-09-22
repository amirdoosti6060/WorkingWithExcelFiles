using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;

Console.WriteLine("Working with Excel file ...");

// License requirement for EPPlus 5 and above
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Define the file path for saving the Excel file
var filePath = new FileInfo("EPPlusExample.xlsx");

if (File.Exists(filePath.FullName))
    File.Delete(filePath.FullName);

using (var package = new ExcelPackage(filePath))
{
    Console.WriteLine("Creating Excel File ...");
    // Create a new worksheet
    var worksheet = package.Workbook.Worksheets.Add("Employee Data");

    // Add some headers
    worksheet.Cells["A1"].Value = "ID";
    worksheet.Cells["B1"].Value = "Name";
    worksheet.Cells["C1"].Value = "Position";
    worksheet.Cells["D1"].Value = "Salary";
    worksheet.Cells["E1"].Value = "Tax";
    worksheet.Cells["F1"].Value = "Net Salary";

    // Apply some styling to the headers
    using (var range = worksheet.Cells["A1:F1"])
    {
        range.Style.Font.Bold = true;
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
    }

    // Add employee data
    var employees = new[]
    {
                new { Id = 1, Name = "John Doe", Position = "Manager", Salary = 5500 },
                new { Id = 2, Name = "Jane Smith", Position = "Engineer", Salary = 4500 },
                new { Id = 3, Name = "Sam Brown", Position = "Technician", Salary = 3000 }
            };

    int row = 2;
    foreach (var employee in employees)
    {
        worksheet.Cells[row, 1].Value = employee.Id;
        worksheet.Cells[row, 2].Value = employee.Name;
        worksheet.Cells[row, 3].Value = employee.Position;
        worksheet.Cells[row, 4].Value = employee.Salary;

        // Add a formula for tax (10% of salary) and net salary
        worksheet.Cells[row, 5].Formula = $"D{row}*0.1";
        worksheet.Cells[row, 6].Formula = $"D{row}-E{row}";

        row++;
    }

    // Apply some currency format to salary, tax, and net salary columns
    using (var range = worksheet.Cells["D2:F4"])
    {
        range.Style.Numberformat.Format = "$#,##0.00";
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
    }

    // AutoFit columns for better display
    worksheet.Cells.AutoFitColumns();

    // Add a chart to represent salary data
    var chart = worksheet.Drawings.AddChart("SalaryChart", eChartType.ColumnClustered);
    chart.Title.Text = "Employee Salary Data";
    chart.SetPosition(6, 0, 1, 0);
    chart.SetSize(600, 300);

    // Add series to the chart
    var series = chart.Series.Add(worksheet.Cells["D2:D4"], worksheet.Cells["B2:B4"]);
    series.Header = "Salary";

    // Save the workbook
    package.Save();

    Console.WriteLine("Excel file created successfully!");
    Console.WriteLine($"File Path: {filePath.FullName}");
}

// Read the data back from the file
using (var package = new ExcelPackage(filePath))
{
    var worksheet = package.Workbook.Worksheets["Employee Data"];
    Console.WriteLine("Reading from the Excel file...");

    for (int row = 2; row <= 4; row++)
    {
        var id = worksheet.Cells[row, 1].Text;
        var name = worksheet.Cells[row, 2].Text;
        var position = worksheet.Cells[row, 3].Text;
        var salary = worksheet.Cells[row, 4].Text;

        Console.WriteLine($"ID: {id}, Name: {name}, Position: {position}, Salary: {salary}");
    }
}
