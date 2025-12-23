using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XSLXtoCSV.Data.UPM_System;
using XSLXtoCSV.Data; // Added for UPMContext
using Microsoft.EntityFrameworkCore;

namespace XSLXtoCSV.Service.Achievement
{
    public static class Assy03_LoadDataService
    {
        
        public static void Normalize(string inputFile, string outputFile, int metadataColumnCount = 5)
        {
            // Usamos ReadLines para no cargar todo en memoria si el archivo es muy grande
            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);

            if (lines.Length < 5)
            {
                throw new InvalidDataException("El archivo de entrada es demasiado corto.");
            }

            // Regex para CSV que respeta comillas
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            // Los headers de fecha están en la fila 2 (índice 2)
            var dateHeaderLine = lines[2];
            var dateHeaders = csvSplitRegex.Split(dateHeaderLine).Skip(metadataColumnCount).ToArray();

            var dataRows = lines.Skip(4); // Los datos empiezan en la fila 4
            var normalizedData = new List<ProductionAchievement>();

            foreach (var row in dataRows)
            {
                if (string.IsNullOrWhiteSpace(row) || row.Trim().Replace(",", "").Length == 0) continue;

                string[] columns = csvSplitRegex.Split(row).Select(s => s.Trim(' ', '"')).ToArray();

                if (columns.Length <= metadataColumnCount) continue;

                // Metadatos básicos
                var currentArea = columns.Length > 1 ? columns[1] : ""; 
                var supervisor = columns.Length > 0 ? columns[0] : ""; 
                var leader = columns.Length > 2 ? columns[2] : ""; 
                var partNumber = columns.Length > 4 ? columns[4] : ""; 

                // Si la fila está vacía en sus columnas clave, saltar
                if (string.IsNullOrWhiteSpace(supervisor) && string.IsNullOrWhiteSpace(partNumber) && string.IsNullOrWhiteSpace(currentArea)) continue;

                // Iterar por bloques de 3 (Tiempo, Objetivo, Real)
                for (int i = 0; i < dateHeaders.Length / 3; i++)
                {
                    int dataIndex = metadataColumnCount + (i * 3);
                    if (dataIndex + 2 >= columns.Length) break;

                    var dateStr = dateHeaders[i * 3];
                    var workingTimeStr = columns[dataIndex];
                    var objectiveStr = columns[dataIndex + 1];
                    var realStr = columns[dataIndex + 2];

                    // Validar que haya datos numéricos antes de procesar
                    bool hasTime = float.TryParse(workingTimeStr, NumberStyles.Any, CultureInfo.InvariantCulture, out float workingTime);
                    bool hasObj = float.TryParse(objectiveStr, NumberStyles.Any, CultureInfo.InvariantCulture, out float objective);
                    bool hasReal = float.TryParse(realStr, NumberStyles.Any, CultureInfo.InvariantCulture, out float real);

                    // Solo agregamos si hay un valor real o un objetivo (evitamos filas vacías)
                    if (!string.IsNullOrWhiteSpace(dateStr) && (real > 0 || objective > 0))
                    {
                        try
                        {
                            var culture = CultureInfo.GetCultureInfo("es-MX");
                            // Limpieza de formato de fecha español
                            var cleanDateStr = dateStr.Replace("a. m.", "AM").Replace("p. m.", "PM").Trim();

                            // Si la fecha solo trae el día, DateTime.ParseExact podría fallar si el formato varía
                            if (DateTime.TryParse(cleanDateStr, culture, DateTimeStyles.None, out DateTime productionDate))
                            {
                                normalizedData.Add(new ProductionAchievement
                                {
                                    Id = Guid.NewGuid(),
                                    Active = true,
                                    CreateDate = DateTime.UtcNow,
                                    CreateBy = "System_Normalize",
                                    ProductionDate = productionDate,
                                    Supervisor = supervisor,
                                    Leader = leader,
                                    Shift = "1",
                                    PartNumberName = $"{partNumber}".Trim(),
                                    WorkingTime = workingTime,
                                    ProductionObjetive = objective,
                                    ProductionReal = real,
                                    Area = currentArea // Dynamically set Area
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log de error silencioso para no detener el proceso
                            Console.WriteLine($"Error en fecha {dateStr}: {ex.Message}");
                        }
                    }
                }
            }

            WriteToCsv(normalizedData, outputFile);
        }

        private static void WriteToCsv(List<ProductionAchievement> data, string filePath)
        {
            var sb = new StringBuilder();
            // Header del CSV
            sb.AppendLine("Id,Active,CreateDate,CreateBy,ProductionDate,Supervisor,Leader,Shift,PartNumberName,WorkingTime,ProductionObjetive,ProductionReal,Area");

            foreach (var r in data)
            {
                // Usamos InvariantCulture para que los decimales sean puntos (.) y no comas (,)
                var line = string.Format(CultureInfo.InvariantCulture,
                    "\"{0}\",\"{1}\",\"{2:s}\",\"{3}\",\"{4:yyyy-MM-dd}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",{9},{10},{11},{12}",
                    r.Id,
                    r.Active,
                    r.CreateDate,
                    r.CreateBy,
                    r.ProductionDate,
                    Sanitize(r.Supervisor),
                    Sanitize(r.Leader),
                    r.Shift,
                    Sanitize(r.PartNumberName),
                    r.WorkingTime,
                    r.ProductionObjetive,
                    r.ProductionReal,
                    r.Area);

                sb.AppendLine(line);
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private static string Sanitize(string field) => field?.Replace("\"", "\"\"") ?? "";

        public static async Task LoadToDatabase(string csvFilePath)
        {
            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine($"Error: CSV file not found at {csvFilePath}");
                return;
            }

            var normalizedData = new List<ProductionAchievement>();
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            try
            {
                var lines = await File.ReadAllLinesAsync(csvFilePath, Encoding.UTF8);

                if (lines.Length <= 1)
                {
                    Console.WriteLine("CSV file is empty or contains only headers.");
                    return;
                }

                foreach (var line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var columns = csvSplitRegex.Split(line).Select(s => s.Trim(' ', '"')).ToArray();

                    try
                    {
                        normalizedData.Add(new ProductionAchievement
                        {
                            Id = Guid.NewGuid(), // Generamos nuevos IDs para la inserción limpia
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.UtcNow,
                            CreateBy = "System_Reload_EF9",
                            ProductionDate = DateTime.Parse(columns[4], CultureInfo.InvariantCulture, DateTimeStyles.None),
                            Supervisor = columns[5],
                            Leader = columns[6],
                            Shift = columns[7],
                            PartNumberName = columns[8],
                            WorkingTime = float.Parse(columns[9], CultureInfo.InvariantCulture),
                            ProductionObjetive = float.Parse(columns[10], CultureInfo.InvariantCulture),
                            ProductionReal = float.Parse(columns[11], CultureInfo.InvariantCulture),
                            Area = columns[12]
                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error parsing line: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing CSV: {ex.Message}");
                return;
            }

            if (normalizedData.Any())
            {
                using (var context = new UPMContext())
                {
                    // Iniciamos una transacción para asegurar la integridad de los datos
                    using var transaction = await context.Database.BeginTransactionAsync();

                    try
                    {
                        // 1. Identificar combinaciones de Área y Periodo (Mes/Año) en el CSV
                        var targets = normalizedData
                            .GroupBy(d => new { d.Area, d.ProductionDate.Year, d.ProductionDate.Month })
                            .Select(g => g.Key);

                        foreach (var target in targets)
                        {
                            // 2. EF9: Ejecutar Delete directo en la DB para limpiar el periodo específico
                            await context.ProductionAchievements
                                .Where(p => p.Area == target.Area
                                         && p.ProductionDate.Year == target.Year
                                         && p.ProductionDate.Month == target.Month)
                                .ExecuteDeleteAsync();

                            Console.WriteLine($"Limpieza completada: Area {target.Area} - Mes {target.Month}/{target.Year}");
                        }

                        // 3. Inserción masiva de los nuevos datos
                        await context.ProductionAchievements.AddRangeAsync(normalizedData);
                        await context.SaveChangesAsync();

                        // 4. Confirmar cambios
                        await transaction.CommitAsync();
                        Console.WriteLine($"Carga completada exitosamente. Total registros: {normalizedData.Count}");
                    }
                    catch (Exception ex)
                    {
                        // Si algo falla, revertimos el borrado
                        await transaction.RollbackAsync();
                        Console.WriteLine($"Error durante la carga. Se realizó Rollback. Detalle: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No se encontraron registros válidos para cargar.");
            }
        }
    }
}


