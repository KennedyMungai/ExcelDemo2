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
