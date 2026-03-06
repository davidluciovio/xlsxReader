using System.Text;
using XSLXtoCSV.Service;
using XSLXtoCSV.Service.Efficiency;



string excelPath = @"";
string destinationPath = $@"";
string month = DateTime.Now.Month.ToString();
string year = DateTime.Now.Year.ToString();

string copiedFile = "";

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

//int area = 5;

//await LoadToDatabase.LoadOperativityToDatabase(@"\\upmap11\c$\UPM\dashboard\operatividad\202512.csv");

await LoadToDatabase.DeleteMontlyData();

for (int area = 1; area <= 5; area++)
{

    switch (area)
    {
        case 1:
            // Ensamble 1
            excelPath = @"\\upms001\USER-ALL\JUNTA DIARIA DE FUNCIONARIOS\2025\02 ENSAMBLE I\CUMPLIMIENTO PIEZAS\2025\2026 APROVECHAMIENTO\3.-APROVECHAMIENTO DIARIO ENS I MAR 26.xlsx";
            destinationPath = $@"\\upmap11\c$\UPM\dashboard\operatividad\ensamble1\{year}{month:00}\";
            break;
        case 2:
            // Ensamble 2
            excelPath = @"\\upms001\USER-ALL\JUNTA DIARIA DE FUNCIONARIOS\2025\02 ENSAMBLE I\CUMPLIMIENTO PIEZAS\2025\2026 APROVECHAMIENTO\3.-APROVECHAMIENTO DIARIO ENS II MAR 26.xlsx";
            destinationPath = $@"\\upmap11\c$\UPM\dashboard\operatividad\ensamble2\{year}{month:00}\";
            break;
        case 3:
            // Estampado
            excelPath = @"\\upms001\USER-ALL\JUNTA DIARIA DE FUNCIONARIOS\2026\19 UPS\% DE APROVECHAMIENTO ESTAMPADO\03. % APROVECHAMIENTO ESTAMPADO - MZO'26 DATOS.xlsx";
            destinationPath = $@"\\upmap11\c$\UPM\dashboard\operatividad\estampado\{year}{month:00}\";
            break;
        case 4:
            // PCP Estampado
            excelPath = @"\\upms002\share\01.Produccion Prensas\2026\Captura\% APROVECHAMIENTO\3.- % APROVECHAMIENTO ESTAMPADO - MARZO.xlsx";
            destinationPath = $@"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\{year}{month:00}\";
            break; 
        case 5:
            // PCP Corte
            excelPath = @"\\upms002\SHARE\01.Producción\05 NUEVO SISTEMA CAPTURA Q,D\CAPTURA PRODUCCION\2026\03 MARZO\CAPTURA DE PRODUCCION MAR.xlsm";
            destinationPath = $@"\\upmap11\c$\UPM\dashboard\operatividad\pcp corte\{year}{month:00}\";
            break;
        default:
            break;
    }


    try
    {
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Error: CSV file not found at {excelPath}");
            return;
        }

        if (!Directory.Exists(destinationPath))
        {
            Directory.CreateDirectory(destinationPath);
            Console.WriteLine($"Directorio creado en: {destinationPath}");
        }

        var fileName = Path.GetFileName(excelPath);
        var destinationFile = Path.Combine(destinationPath, fileName);

        File.Copy(excelPath, destinationFile, true);

        copiedFile = destinationFile;

        Console.WriteLine($"Archivo {excelPath} copiado y reemplazado (si existía).");

    }
    catch (Exception ex)
    {
        Console.WriteLine($"ERROR AL COPIAR EL ARCHIVO: '{excelPath}', ERROR: '{ex.Message}'");
    }



    try
    {
        var csvFiles = new List<string>();
        var csvNormalizeFiles = new List<string>();

        ConvertSheetsToCSV.ProcessExcelFixed(copiedFile, ref csvFiles);
        Console.WriteLine("\n¡Proceso Ensamble finalizado correctamente!");

        csvFiles.ForEach(async file =>
        {
            switch (area)
            {
                case 1:
                    Assy01_OperationalEfficiency_LoadData.NormalizeEfficiency(file, file.Replace(".csv", "_Normalize.csv"));
                    Console.WriteLine($"\nNormalizar Ensamble ¡Proceso finalizado correctamente para el archivo: {file}!");
                    break;
                case 2:
                    Assy02_OperationalEfficiency_LoadData.NormalizeEfficiencyEnsII(file, file.Replace(".csv", "_Normalize.csv"));
                    Console.WriteLine($"\nNormalizar Ensamble ¡Proceso finalizado correctamente para el archivo: {file}!");
                    break;
                case 3:
                    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(file, file.Replace(".csv", "_Normalize.csv"));
                    Console.WriteLine($"\nNormalizar Estampado ¡Proceso finalizado correctamente para el archivo: {file}!");
                    break;
                case 4:
                    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(file, file.Replace(".csv", "_Normalize.csv"));
                    Console.WriteLine($"\nNormalizar PCP Estampado ¡Proceso finalizado correctamente para el archivo: {file}!");
                    break;
                case 5:
                    PCPCorte_OperationalEfficiency_LoadData.NormalizeCorte(file, file.Replace(".csv", "_Normalize.csv"));
                    Console.WriteLine($"\nNormalizar PCP Corte ¡Proceso finalizado correctamente para el archivo: {file}!");
                    break;
            }

            csvNormalizeFiles.Add(file.Replace(".csv", "_Normalize.csv"));
        });

        foreach (var file in csvNormalizeFiles)
        {
            await LoadToDatabase.LoadOperativityToDatabase(file);
            Console.WriteLine($"\nCargar datos de Operatividad desde {file} ¡Proceso finalizado correctamente!");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.Message);
    }
}


#region all data paths 


///******************************************************************************/
////Original Excel Files Paths

//string excelAssyPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx";
//string excelStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx";
//string excelCortePath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\corte\% CUMPLIMIENTO T. TACTO - TM.xlsx";

//string excelAssy01Path_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx";
//string excelAssy02Path_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx";

//string excelStampPath_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx";

//string excelPCPStampPath_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx";

//string excelPCPCortePath_OperationalEfficiency = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp corte\% CUMPLIMIENTO T. TACTO - TM+PIEZAS.xlsx";


///******************************************************************************/
////Achievement CSV Files Paths
////Assy Files Paths
//string _fileAssy01Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE I.csv";
//string _fileAssy01OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE I_Normalize.csv";

//string _fileAssy02Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE II.csv";
//string _fileAssy02OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE II_Normalize.csv";

//string _fileAssy03Path = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE III.csv";
//string _fileAssy03OutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\ensamble\Copia de PIEZAS DIARIAS DICIEMBRE.xlsx_ENSAMBLE III.csv_Normalize.csv";

////Corte Files Paths
//string _fileCortePath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\corte\% CUMPLIMIENTO T. TACTO - TM.xlsx_CORTE Y ENSAMBLE.csv";
//string _fileCorteOutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\corte\% CUMPLIMIENTO T. TACTO - TM.xlsx_CORTE Y ENSAMBLE_Normalize.csv";

////Stamp Files Paths
//string _fileStampPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx_SPM.csv";
//string _fileStampOutPath = @"\\upmap11\c$\UPM\dashboard\cumplimiento\estampado\Control de resultados por grupos-Diciembre 25.xlsx_SPM.csv_Normalize.csv";

///******************************************************************************/
////operational Efficiency CSV Files Paths
////Assy 01 Operational Efficiency Files Paths
//string _fileAssy01OEShift01Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT..csv";
//string _fileAssy01OEShift01OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT_Normalize.csv";

//string _fileAssy01OEShift03Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT..csv";
//string _fileAssy01OEShift03OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS I DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT_Normalize.csv";

////Assy 02 Operational Efficiency Files Paths
//string _fileAssy02OEShift01Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT..csv";
//string _fileAssy02OEShift01OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 1ERT_Normalize.csv";

//string _fileAssy02OEShift03Path = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT.csv";
//string _fileAssy02OEShift03OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\ensamble\APROVECHAMIENTO DIARIO ENS II DIC 25.xlsx_APROVECHAMIENTO DIARIO 3ERT_Normalize.csv";

////Stamp Operational Efficiency Files Paths
//string _fileStampOEBLK600Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 600.csv";
//string _fileStampOEBLK600OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 600_Normalize.csv";

//string _fileStampOEBLK800Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 800.csv";
//string _fileStampOEBLK800OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK 800_Normalize.csv";

//string _fileStampOEBLKIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK I.csv";
//string _fileStampOEBLKIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK I_Normalize.csv";

//string _fileStampOEBLKIIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK II.csv";
//string _fileStampOEBLKIIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_BLK II_Normalize.csv";

//string _fileStampOELASERIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER I.csv";
//string _fileStampOELASERIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER I_Normalize.csv";

//string _fileStampOELASERIIIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER III.csv";
//string _fileStampOELASERIIIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_LASER III_Normalize.csv";

//string _fileStampOETNDPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TND.csv";
//string _fileStampOETNDOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TND_Normalize.csv";

//string _fileStampOETRF1500Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 1500.csv";
//string _fileStampOETRF1500OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 1500_Normalize.csv";

//string _fileStampOETRF2000Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2000.csv";
//string _fileStampOETRF2000OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2000_Normalize.csv";

//string _fileStampOETRF2500IIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500 II.csv";
//string _fileStampOETRF2500IIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500 II_Normalize.csv";

//string _fileStampOETRF2500Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500.csv";
//string _fileStampOETRF2500OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 2500_Normalize.csv";

//string _fileStampOETRF3000Path = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 3000.csv";
//string _fileStampOETRF3000OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF 3000_Normalize.csv";

//string _fileStampOETRFIIIPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF III.csv";
//string _fileStampOETRFIIIOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE'25 DATOS.xlsx_TRF III_Normalize.csv";

////PCP Stamp Operational Efficiency Files Paths
//string _filePCPStampOETRF800Path = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TRF 800.csv";
//string _filePCPStampOETRF800OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TRF 800_Normalize.csv";

//string _filePCPStampOETRF1200Path = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TRF 1200.csv";
//string _filePCPStampOETRF1200OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TRF 1200_Normalize.csv";

//string _filePCPStampOEBLK400Path = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_BLK400.csv";
//string _filePCPStampOEBLK400OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_BLK400_Normalize.csv";

//string _filePCPStampOETandem2Path = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TANDEM 2.csv";
//string _filePCPStampOETandem2OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TANDEM 2_Normalize.csv";

//string _filePCPStampOETandemPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TANDEM.csv";
//string _filePCPStampOETandemOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TANDEM_Normalize.csv";

//string _filePCPStampOETRF630Path = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TRF 630.csv";
//string _filePCPStampOETRF630OutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp estampado\% APROVECHAMIENTO ESTAMPADO - DICIEMBRE.xlsx_TRF 630_Normalize.csv";

////PCP Corte Operational Efficiency Files Paths
//string _filePCPCortePath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp corte\% CUMPLIMIENTO T. TACTO - TM+PIEZAS.xlsx_CORTE Y ENSAMBLE.csv";
//string _filePCPCorteOutPath = @"\\upmap11\c$\UPM\dashboard\operatividad\pcp corte\% CUMPLIMIENTO T. TACTO - TM+PIEZAS.xlsx_CORTE Y ENSAMBLE_Normalize.csv";

///******************************************************************************/


////Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

//try
//{
//    ConvertSheetsToCSV.ProcessExcelFixed(excelAssyPath);
//    Console.WriteLine("\n¡Proceso Ensamble finalizado correctamente!");

//    ConvertSheetsToCSV.ProcessExcelFixed(excelStampPath);
//    Console.WriteLine("\n¡Proceso Estampado finalizado correctamente!");

//    ConvertSheetsToCSV.ProcessExcelFixed(excelCortePath);
//    Console.WriteLine("\n¡Proceso Corte finalizado correctamente!");

//    /******************************************************************************/

//    ConvertSheetsToCSV.ProcessExcelFixed(excelAssy01Path_OperationalEfficiency);
//    Console.WriteLine("\n¡Proceso Ensamble I Operatividad finalizado correctamente!");

//    ConvertSheetsToCSV.ProcessExcelFixed(excelAssy02Path_OperationalEfficiency);
//    Console.WriteLine("\n¡Proceso Ensamble II Operatividad finalizado correctamente!");

//    ConvertSheetsToCSV.ProcessExcelFixed(excelStampPath_OperationalEfficiency);
//    Console.WriteLine("\n¡Proceso Estampado Operatividad finalizado correctamente!");

//    ConvertSheetsToCSV.ProcessExcelFixed(excelPCPStampPath_OperationalEfficiency);
//    Console.WriteLine("\n¡Proceso PCP Estampado Operatividad finalizado correctamente!");

//    /******************************************************************************/

//    ConvertSheetsToCSV.ProcessExcelFixed(excelPCPCortePath_OperationalEfficiency);
//    Console.WriteLine("\n¡Proceso PCP Corte Operatividad finalizado correctamente!");

//    /******************************************************************************/

//    Corte_LoadDataService.Normalize(_fileCortePath, _fileCorteOutPath);
//    Console.WriteLine("\nNormalizar Corte ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    Assy01_LoadDataService.Normalize(_fileAssy01Path, _fileAssy01OutPath);
//    Console.WriteLine("\nNormalizar Ensamble 01 ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    Assy01_OperationalEfficiency_LoadData.NormalizeEfficiency(_fileAssy01OEShift01Path, _fileAssy01OEShift01OutPath);
//    Console.WriteLine("\nNormalizar Ensamble 01 turno 1 Operatividad ¡Proceso finalizado correctamente!");

//    Assy01_OperationalEfficiency_LoadData.NormalizeEfficiency(_fileAssy01OEShift03Path, _fileAssy01OEShift03OutPath, "ENSAMBLE I", "3");
//    Console.WriteLine("\nNormalizar Ensamble 01 turno 3 Operatividad ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    Assy02_LoadDataService.Normalize(_fileAssy02Path, _fileAssy02OutPath);
//    Console.WriteLine("\nNormalizar Ensamble 02 ¡Proceso finalizado correctamente!");


//    Assy02_OperationalEfficiency_LoadData.NormalizeEfficiencyEnsII(_fileAssy02OEShift01Path, _fileAssy02OEShift01OutPath);
//    Console.WriteLine("\nNormalizar Ensamble 02 turno 1 Operatividad ¡Proceso finalizado correctamente!");

//    Assy02_OperationalEfficiency_LoadData.NormalizeEfficiencyEnsII(_fileAssy02OEShift03Path, _fileAssy02OEShift03OutPath);
//    Console.WriteLine("\nNormalizar Ensamble 01 turno 3 Operatividad ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    Assy03_LoadDataService.Normalize(_fileAssy03Path, _fileAssy03OutPath);
//    Console.WriteLine("\nNormalizar Ensamble 03 ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    Stamp_LoadDataService.NormalizeSPM(_fileStampPath, _fileStampOutPath);
//    Console.WriteLine("\nNormalizar Estampado ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF2500Path, _fileStampOETRF2500OutPath);
//    Console.WriteLine("\nNormalizar Estampado TRF 2500 ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF2500IIPath, _fileStampOETRF2500IIOutPath);
//    Console.WriteLine("\nNormalizar Estampado TRF 2500 II ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF2000Path, _fileStampOETRF2000OutPath);
//    Console.WriteLine("\nNormalizar Estampado TRF 2000 ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF1500Path, _fileStampOETRF1500OutPath);
//    Console.WriteLine("\nNormalizar Estampado TRF 1500 ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETNDPath, _fileStampOETNDOutPath);
//    Console.WriteLine("\nNormalizar Estampado TND ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOELASERIIIPath, _fileStampOELASERIIIOutPath);
//    Console.WriteLine("\nNormalizar Estampado LASER III ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOELASERIPath, _fileStampOELASERIOutPath);
//    Console.WriteLine("\nNormalizar Estampado LASER I ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLKIIPath, _fileStampOEBLKIIOutPath);
//    Console.WriteLine("\nNormalizar Estampado BLK II ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLKIPath, _fileStampOEBLKIOutPath);
//    Console.WriteLine("\nNormalizar Estampado BLK I ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLK800Path, _fileStampOEBLK800OutPath);
//    Console.WriteLine("\nNormalizar Estampado BLK 800 ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRF3000Path, _fileStampOETRF3000OutPath);
//    Console.WriteLine("\nNormalizar Estampado TRF 3000 ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOEBLK600Path, _fileStampOEBLK600OutPath);
//    Console.WriteLine("\nNormalizar Estampado BLK 600 ¡Proceso finalizado correctamente!");

//    Stamp_OperationalEfficiency_LoadData.NormalizeEstampado(_fileStampOETRFIIIPath, _fileStampOETRFIIIOutPath);
//    Console.WriteLine("\nNormalizar Estampado TRF III ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(_filePCPStampOETRF800Path, _filePCPStampOETRF800OutPath);
//    Console.WriteLine("\nNormalizar PCP Estampado TRF 800 ¡Proceso finalizado correctamente!");

//    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(_filePCPStampOETRF1200Path, _filePCPStampOETRF1200OutPath);
//    Console.WriteLine("\nNormalizar PCP Estampado TRF 1200 ¡Proceso finalizado correctamente!");

//    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(_filePCPStampOEBLK400Path, _filePCPStampOEBLK400OutPath);
//    Console.WriteLine("\nNormalizar PCP Estampado BLK 400 ¡Proceso finalizado correctamente!");

//    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(_filePCPStampOETandem2Path, _filePCPStampOETandem2OutPath);
//    Console.WriteLine("\nNormalizar PCP Estampado TANDEM 2 ¡Proceso finalizado correctamente!");

//    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(_filePCPStampOETandemPath, _filePCPStampOETandemOutPath);
//    Console.WriteLine("\nNormalizar PCP Estampado TANDEM ¡Proceso finalizado correctamente!");

//    PCPStamp_OperationalEfficiency_LoadData.NormalizeEstampado(_filePCPStampOETRF630Path, _filePCPStampOETRF630OutPath);
//    Console.WriteLine("\nNormalizar PCP Estampado TRF 630 ¡Proceso finalizado correctamente!");

//    /******************************************************************************/

//    PCPCorte_OperationalEfficiency_LoadData.NormalizeCorte(_filePCPCortePath, _filePCPCorteOutPath);
//    Console.WriteLine("\nNormalizar PCP Corte ¡Proceso finalizado correctamente!");


//}
//catch (Exception ex)
//{
//    Console.WriteLine($"Error: {ex.Message}");
//}



//try
//{
//    List<string> AchievementFiles = new List<string>
//    {
//        _fileAssy01OutPath,
//        _fileAssy02OutPath,
//        _fileAssy03OutPath,
//        _fileCorteOutPath,
//        _fileStampOutPath
//    };

//    foreach (var file in AchievementFiles)
//    {
//        await LoadToDatabase.LoadAchievementToDatabase(file);
//        Console.WriteLine($"\nCargar datos de Cumplimiento desde {file} ¡Proceso finalizado correctamente!");
//    }

//    /******************************************************************************/

//    List<string> EfficiencyFiles = new List<string>
//    {
//        _fileAssy01OEShift01OutPath,
//        _fileAssy01OEShift03OutPath,

//        _fileAssy02OEShift01OutPath,
//        _fileAssy02OEShift03OutPath,

//        _fileStampOEBLK600OutPath,
//        _fileStampOEBLK800OutPath,
//        _fileStampOEBLKIOutPath,
//        _fileStampOEBLKIIOutPath,
//        _fileStampOELASERIOutPath,
//        _fileStampOELASERIIIOutPath,
//        _fileStampOETNDOutPath,
//        _fileStampOETRF1500OutPath,
//        _fileStampOETRF2000OutPath,
//        _fileStampOETRF2500IIOutPath,
//        _fileStampOETRF2500OutPath,
//        _fileStampOETRF3000OutPath,
//        _fileStampOETRFIIIOutPath,

//        _filePCPStampOETRF800OutPath,
//        _filePCPStampOETRF1200OutPath,
//        _filePCPStampOEBLK400OutPath,
//        _filePCPStampOETandem2OutPath,
//        _filePCPStampOETandemOutPath,
//        _filePCPStampOETRF630OutPath,

//        _filePCPCorteOutPath
//    };

//    foreach (var file in EfficiencyFiles)
//    {
//        await LoadToDatabase.LoadOperativityToDatabase(file);
//        Console.WriteLine($"\nCargar datos de Operatividad desde {file} ¡Proceso finalizado correctamente!");
//    }

//}
//catch (Exception ex)
//{
//    Console.WriteLine($"Error: {ex.Message}");
//}

#endregion
