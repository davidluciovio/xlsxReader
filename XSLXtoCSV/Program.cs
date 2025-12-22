using ExcelDataReader;
using System.Data;
using System.Text;
using XSLXtoCSV.Service;
using XSLXtoCSV.Service.Achievement;
using XSLXtoCSV.Service.Efficiency;


string excelAssyPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx";
string excelStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx";

string excelAssy01Path_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx";
string excelAssy02Path_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx";

string excelStampPath_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25.xlsx";

/******************************************************************************/

string _fileAssy01Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE I.csv";
string _fileAssy01OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE I_Normalize.csv";

string _fileAssy01OEShift01Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT..csv";
string _fileAssy01OEShift01OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT_Normalize.csv";

string _fileAssy01OEShift03Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT..csv";
string _fileAssy01OEShift03OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT_Normalize.csv";

/******************************************************************************/

string _fileAssy02Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE II.csv";
string _fileAssy02OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE II_Normalize.csv";

string _fileAssy02OEShift01Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT..csv";
string _fileAssy02OEShift01OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT_Normalize.csv";

string _fileAssy02OEShift03Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT.csv";
string _fileAssy02OEShift03OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT_Normalize.csv";

/******************************************************************************/

string _fileAssy03Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE III.csv";
string _fileAssy03OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE III.csv_Normalize.csv";

/******************************************************************************/

string _fileStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx_SPM.csv";
string _fileStampOutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx_SPM.csv_Normalize.csv";

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

try
{
    ConvertSheetsToCSV.ProcessExcelFixed(excelAssyPath);
    Console.WriteLine("\n¡Proceso Ensamble finalizado correctamente!");

    ConvertSheetsToCSV.ProcessExcelFixed(excelStampPath);
    Console.WriteLine("\n¡Proceso Estampado finalizado correctamente!");

    /******************************************************************************/

    ConvertSheetsToCSV.ProcessExcelFixed(excelAssy01Path_OperationalEfficiency);
    Console.WriteLine("\n¡Proceso Ensamble I Operatividad finalizado correctamente!");

    ConvertSheetsToCSV.ProcessExcelFixed(excelAssy02Path_OperationalEfficiency);
    Console.WriteLine("\n¡Proceso Ensamble II Operatividad finalizado correctamente!");

    ConvertSheetsToCSV.ProcessExcelFixed(excelStampPath_OperationalEfficiency);
    Console.WriteLine("\n¡Proceso Estampado Operatividad finalizado correctamente!");

    /******************************************************************************/

    Assy01_LoadDataService.Normalize(_fileAssy01Path, _fileAssy01OutPath);
    Console.WriteLine("\nNormalizar Ensamble 01 ¡Proceso finalizado correctamente!");


    Assy01_OperationalEfficiency_LoadData.NormalizeEfficiency(_fileAssy01OEShift01Path, _fileAssy01OEShift01OutPath);
    Console.WriteLine("\nNormalizar Ensamble 01 turno 1 Operatividad ¡Proceso finalizado correctamente!");

    Assy01_OperationalEfficiency_LoadData.NormalizeEfficiency(_fileAssy01OEShift03Path, _fileAssy01OEShift03OutPath, "ENSAMBLE I", "3");
    Console.WriteLine("\nNormalizar Ensamble 01 turno 3 Operatividad ¡Proceso finalizado correctamente!");

    /******************************************************************************/

    Assy02_LoadDataService.Normalize(_fileAssy02Path, _fileAssy02OutPath);
    Console.WriteLine("\nNormalizar Ensamble 02 ¡Proceso finalizado correctamente!");


    Assy02_OperationalEfficiency_LoadData.NormalizeEfficiencyEnsII(_fileAssy02OEShift01Path, _fileAssy02OEShift01OutPath);
    Console.WriteLine("\nNormalizar Ensamble 02 turno 1 Operatividad ¡Proceso finalizado correctamente!");

    Assy02_OperationalEfficiency_LoadData.NormalizeEfficiencyEnsII(_fileAssy02OEShift03Path, _fileAssy02OEShift03OutPath);
    Console.WriteLine("\nNormalizar Ensamble 01 turno 3 Operatividad ¡Proceso finalizado correctamente!");

    /******************************************************************************/

    Assy03_LoadDataService.Normalize(_fileAssy03Path, _fileAssy03OutPath);
    Console.WriteLine("\nNormalizar Ensamble 03 ¡Proceso finalizado correctamente!");

    /******************************************************************************/

    Stamp_LoadDataService.NormalizeSPM(_fileStampPath, _fileStampOutPath);
    Console.WriteLine("\nNormalizar Estampado ¡Proceso finalizado correctamente!");

    /******************************************************************************/

    await Assy01_LoadDataService.LoadToDatabase(_fileAssy01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 ¡Proceso finalizado correctamente!");

    await Assy01_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy01OEShift01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 1 Operatividad  ¡Proceso finalizado correctamente!");

    await Assy01_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy01OEShift03OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 3 Operatividad  ¡Proceso finalizado correctamente!");

    await Assy02_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy02OEShift01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 1 Operatividad  ¡Proceso finalizado correctamente!");

    await Assy02_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy02OEShift03OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 3 Operatividad  ¡Proceso finalizado correctamente!");
    
    await Assy02_LoadDataService.LoadToDatabase(_fileAssy02OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 02 ¡Proceso finalizado correctamente!");
    
    await Stamp_LoadDataService.LoadToDatabase(_fileStampOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado ¡Proceso finalizado correctamente!");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}


    
