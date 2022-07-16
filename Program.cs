using System.Drawing;
using System;
using System.ComponentModel;
using OfficeOpenXml;
using System.IO;
using static System.Environment;

namespace ExcelDemo2;

public class Program
{
    public static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        var file = new FileInfo(Path.Combine(SpecialFolder.Desktop + "ExcelSucks.xlsx"));

        var people = GetSetupData();

        await SaveExcelFile(people, file);
    }

    private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
    {
        DeleteIfExists(file);

        using var package = new ExcelPackage(file);

        var ws = package.Workbook.Worksheets.Add("MainReport");
        
        var range = ws.Cells["A2"].LoadFromCollection<PersonModel>(people, true);
        range.AutoFitColumns();

        // Formatting for the header row
        ws.Cells["A1"].Value = "Our cool Report";
        ws.Cells["A1:C1"].Merge = true;
        ws.Column(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
        ws.Row(1).Style.Font.Size = 24;
        ws.Row(1).Style.Font.Color.SetColor(Color.Blue);

        await package.SaveAsync();
    }

    private static void DeleteIfExists(FileInfo file)
    {
        if(file.Exists)
        {
            file.Delete();
        }
    }

    static List<PersonModel> GetSetupData()
    {
        List<PersonModel> output = new()
        {
            new() {Id=1, FirstName="The", LastName="Snitch"},
            new() {Id=1, FirstName="is", LastName="Out"},
            new() {Id=1, FirstName="Of", LastName="Control"}
        };

        return output;
    }
}
