using ExcelDataReader;
using System.Data;
using System.Text;
using XSLXtoCSV.Service;
using XSLXtoCSV.Service.Achievement;
using XSLXtoCSV.Service.Efficiency;


string excelAssyPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx";
string excelStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx";

string excelAssy01Path_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx";
string excelAssy02Path_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx";

string excelStampPath_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx";

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

string _fileStampOEBLK600Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 600.csv";
string _fileStampOEBLK600OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 600_Normalize.csv";

string _fileStampOEBLK800Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 800.csv";
string _fileStampOEBLK800OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 800_Normalize.csv";

string _fileStampOEBLKIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK I.csv";
string _fileStampOEBLKIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK I_Normalize.csv";

string _fileStampOEBLKIIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK II.csv";
string _fileStampOEBLKIIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK II_Normalize.csv";

string _fileStampOELASERIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER I.csv";
string _fileStampOELASERIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER I_Normalize.csv";

string _fileStampOELASERIIIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER III.csv";
string _fileStampOELASERIIIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER III_Normalize.csv";

string _fileStampOETNDPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TND.csv";
string _fileStampOETNDOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TND_Normalize.csv";

string _fileStampOETRF1500Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 1500.csv";
string _fileStampOETRF1500OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 1500_Normalize.csv";

string _fileStampOETRF2000Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2000.csv";
string _fileStampOETRF2000OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2000_Normalize.csv";

string _fileStampOETRF2500IIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500 II.csv";
string _fileStampOETRF2500IIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500 II_Normalize.csv";

string _fileStampOETRF2500Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500.csv";
string _fileStampOETRF2500OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500_Normalize.csv";

string _fileStampOETRF3000Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 3000.csv";
string _fileStampOETRF3000OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 3000_Normalize.csv";

string _fileStampOETRFIIIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF III.csv";
string _fileStampOETRFIIIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 3000_Normalize.csv";

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

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF2500Path, _fileStampOETRF2500OutPath);
    Console.WriteLine("\nNormalizar Estampado TRF 2500 ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF2500IIPath, _fileStampOETRF2500IIOutPath);
    Console.WriteLine("\nNormalizar Estampado TRF 2500 II ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF2000Path, _fileStampOETRF2000OutPath);
    Console.WriteLine("\nNormalizar Estampado TRF 2000 ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF1500Path, _fileStampOETRF1500OutPath);
    Console.WriteLine("\nNormalizar Estampado TRF 1500 ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETNDPath, _fileStampOETNDOutPath);
    Console.WriteLine("\nNormalizar Estampado TND ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOELASERIIIPath, _fileStampOELASERIIIOutPath);
    Console.WriteLine("\nNormalizar Estampado LASER III ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOELASERIPath, _fileStampOELASERIOutPath);
    Console.WriteLine("\nNormalizar Estampado LASER I ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLKIIPath, _fileStampOEBLKIIOutPath);
    Console.WriteLine("\nNormalizar Estampado BLK II ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLKIPath, _fileStampOEBLKIOutPath);
    Console.WriteLine("\nNormalizar Estampado BLK I ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLK800Path, _fileStampOEBLK800OutPath);
    Console.WriteLine("\nNormalizar Estampado BLK 800 ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF3000Path, _fileStampOETRF3000OutPath);
    Console.WriteLine("\nNormalizar Estampado TRF 3000 ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLK600Path, _fileStampOEBLK600OutPath);
    Console.WriteLine("\nNormalizar Estampado BLK 600 ¡Proceso finalizado correctamente!");

    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRFIIIPath, _fileStampOETRFIIIOutPath);
    Console.WriteLine("\nNormalizar Estampado TRF III ¡Proceso finalizado correctamente!");

    /******************************************************************************/

    await Assy01_LoadDataService.LoadToDatabase(_fileAssy01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 ¡Proceso finalizado correctamente!");
    
    await Assy02_LoadDataService.LoadToDatabase(_fileAssy02OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 02 ¡Proceso finalizado correctamente!");

    await Assy01_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy01OEShift01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 1 Operatividad  ¡Proceso finalizado correctamente!");
    await Assy01_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy01OEShift03OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 3 Operatividad  ¡Proceso finalizado correctamente!");

    await Assy02_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy02OEShift01OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 1 Operatividad  ¡Proceso finalizado correctamente!");
    await Assy02_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileAssy02OEShift03OutPath);
    Console.WriteLine("\nInsertar en base de datos Ensamble 01 turno 3 Operatividad  ¡Proceso finalizado correctamente!");
    
    await Stamp_LoadDataService.LoadToDatabase(_fileStampOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado ¡Proceso finalizado correctamente!");

    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETRF2500OutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TRF 2500 ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETRF2500IIOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TRF 2500 II ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETRF2000OutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TRF 2000 ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETRF1500OutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TRF 1500 ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETNDOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TND ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOELASERIIIOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado LASER III ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOELASERIOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado LASER I ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOEBLKIIOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado BLK II ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOEBLKIOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado BLK I ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOEBLK800OutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado BLK 800 ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETRF3000OutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TRF 3000 ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOEBLK600OutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado BLK 600 ¡Proceso finalizado correctamente!");
    await Stamp_OperationalEfficiency_LoadData.LoadEfficiencyToDatabase(_fileStampOETRFIIIOutPath);
    Console.WriteLine("\nInsertar en base de datos Estampado TRF III ¡Proceso finalizado correctamente!");


}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}


    
