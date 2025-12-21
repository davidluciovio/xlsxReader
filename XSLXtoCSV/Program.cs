using ExcelDataReader;
using System.Data;
using System.Text;
using XSLXtoCSV.Service;
using XSLXtoCSV.Service.Achievement;


string excelAssyPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx";
string excelStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx";

string _fileAssy01Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE I.csv";
string _fileAssy01OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE I_Normalize.csv";

string _fileAssy02Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE II.csv";
string _fileAssy02OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE II_Normalize.csv";

string _fileAssy03Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE III.csv";
string _fileAssy03OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE III.csv_Normalize.csv";

string _fileStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx_SPM.csv";
string _fileStampOutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx_SPM.csv_Normalize.csv";

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

try
{
    ConvertSheetsToCSV.ProcessExcelFixed(excelAssyPath);
    Console.WriteLine("\n¡Proceso Ensamble finalizado correctamente!");

    ConvertSheetsToCSV.ProcessExcelFixed(excelStampPath);
    Console.WriteLine("\n¡Proceso Estampado finalizado correctamente!");


    Assy01_LoadDataService.Normalize(_fileAssy01Path, _fileAssy01OutPath);
    Console.WriteLine("\nNormalizar Ensamble 01 ¡Proceso finalizado correctamente!");

    await Assy01_LoadDataService.LoadToDatabase(_fileAssy01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 ¡Proceso finalizado correctamente!");


    Assy02_LoadDataService.Normalize(_fileAssy02Path, _fileAssy02OutPath);
    Console.WriteLine("\nNormalizar Ensamble 02 ¡Proceso finalizado correctamente!");

    await Assy02_LoadDataService.LoadToDatabase(_fileAssy02OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 02 ¡Proceso finalizado correctamente!");


    Assy03_LoadDataService.Normalize(_fileAssy03Path, _fileAssy03OutPath);
    Console.WriteLine("\nNormalizar Ensamble 03 ¡Proceso finalizado correctamente!");


    Stamp_LoadDataService.NormalizeSPM(_fileStampPath, _fileStampOutPath);
    Console.WriteLine("\nNormalizar Estampado ¡Proceso finalizado correctamente!");

    await Stamp_LoadDataService.LoadToDatabase(_fileStampOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado ¡Proceso finalizado correctamente!");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}


    
