using System;
using System.ComponentModel;
using OfficeOpenXml;
using System.IO;
using static System.Environment;

namespace ExcelDemo2;

public class Program
{
    public static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

        var file = new FileInfo(SpecialFolder.Desktop.ToString());
    }
}