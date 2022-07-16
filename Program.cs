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

    private static Task SaveExcelFile(List<PersonModel> people, FileInfo file)
    {
        DeleteIfExists(file);

        using var package = new ExcelPackage(file);
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
