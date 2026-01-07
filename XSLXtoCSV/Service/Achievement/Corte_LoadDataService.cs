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
    public static class Corte_LoadDataService
    {

        public static void Normalize(string inputFile, string outputFile)
        {
            // Leemos el archivo completo para procesar celdas multilínea correctamente
            string fullText = File.ReadAllText(inputFile, Encoding.UTF8);

            // Regex para separar filas y columnas respetando comillas
            var rowSplitRegex = new Regex(@"\r?\n(?=(?:[^""]*""[^""]*"")*[^""]*$)");
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            var lines = rowSplitRegex.Split(fullText).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();

            // Variables de contexto que se actualizan en cada bloque de encabezados
            string[] currentDates = null;
            string[] currentSupervisors = null;
            string[] currentLeaders = null;
            string[] currentShifts = null;

            int metadataColumnCount = 6;
            var normalizedData = new List<ProductionAchievement>();

            for (int i = 0; i < lines.Length; i++)
            {
                var columns = csvSplitRegex.Split(lines[i]).Select(s => s.Trim(' ', '"', '\n', '\r')).ToArray();

                // 1. DETECCIÓN DE ENCABEZADOS (Reinicio de contexto)
                // Buscamos la fila que identifica el inicio de una sección
                if (columns.Length > 2 && columns[0] == "TIEMPO TACTO HP" && columns[1] == "ITEM" && columns[2] == "LINEA")
                {
                    // Fila actual: Fechas
                    currentDates = columns;

                    // Fila +1: Supervisores
                    if (i + 1 < lines.Length)
                        currentSupervisors = csvSplitRegex.Split(lines[i + 1]).Select(s => s.Trim(' ', '"', '\n', '\r')).ToArray();

                    // Fila +2: Líderes (Limpiamos saltos de línea internos)
                    if (i + 2 < lines.Length)
                        currentLeaders = csvSplitRegex.Split(lines[i + 2]).Select(s => s.Trim(' ', '"').Replace("\n", " / ")).ToArray();

                    // Fila +4: Turnos (Una fila debajo del líder, saltando la repetición del líder en la fila 3)
                    if (i + 4 < lines.Length)
                    {
                        currentShifts = csvSplitRegex.Split(lines[i + 4]).Select(s => {
                            var val = s.Trim(' ', '"', '\n', '\r');
                            var match = Regex.Match(val, @"\d+"); // Extrae "1" de "1er"
                            return match.Success ? match.Value : "1";
                        }).ToArray();
                    }

                    // Saltamos el bloque completo de encabezados (6 filas)
                    i += 5;
                    continue;
                }

                // 2. PROCESAMIENTO DE DATOS
                // Verificamos que sea una fila de producción (Columna ITEM es numérica)
                if (currentDates != null && columns.Length > metadataColumnCount && int.TryParse(columns[1], out _))
                {
                    var partName = columns[2]; // LINEA
                    float.TryParse(columns[4], NumberStyles.Any, CultureInfo.InvariantCulture, out float targetObjective);

                    // Cada día tiene 4 columnas, procesamos los dos sub-bloques (2 grupos de líderes/turnos)
                    for (int dayIdx = 0; dayIdx < (currentDates.Length - metadataColumnCount) / 4; dayIdx++)
                    {
                        for (int subIdx = 0; subIdx < 2; subIdx++)
                        {
                            int colIdx = metadataColumnCount + (dayIdx * 4) + (subIdx * 2);
                            if (colIdx + 1 >= columns.Length) break;

                            float.TryParse(columns[colIdx], NumberStyles.Any, CultureInfo.InvariantCulture, out float tactTimeReal);
                            float.TryParse(columns[colIdx + 1], NumberStyles.Any, CultureInfo.InvariantCulture, out float compliancePercent);

                            // Solo guardamos si hay actividad reportada
                            if (tactTimeReal > 0 || compliancePercent > 0)
                            {
                                var dateStr = currentDates[colIdx];
                                var culture = CultureInfo.GetCultureInfo("es-MX");
                                var cleanDate = dateStr.Replace("a. m.", "AM").Replace("p. m.", "PM").Trim();

                                if (DateTime.TryParse(cleanDate, culture, DateTimeStyles.None, out DateTime prodDate))
                                {
                                    normalizedData.Add(new ProductionAchievement
                                    {
                                        Id = Guid.NewGuid(),
                                        Active = true,
                                        CreateDate = DateTime.UtcNow,
                                        CreateBy = "System_Normalize_TactTime_V2",
                                        ProductionDate = prodDate,
                                        // Asignamos valores desde el contexto actual
                                        Supervisor = (colIdx < currentSupervisors?.Length) ? currentSupervisors[colIdx] : "S/S",
                                        Leader = (colIdx < currentLeaders?.Length) ? currentLeaders[colIdx] : "S/L",
                                        Shift = (colIdx < currentShifts?.Length) ? currentShifts[colIdx] : "1",
                                        PartNumberName = partName,
                                        WorkingTime = tactTimeReal,
                                        ProductionObjetive = targetObjective,
                                        ProductionReal = compliancePercent,
                                        Area = "CORTE PCP"
                                    });
                                }
                            }
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
                    Console.WriteLine("CSV file is empty or contains only headers. No data to load.");
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
                            Id = Guid.Parse(columns[0]),
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.Parse(columns[2], CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal),
                            CreateBy = columns[3],
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
                    catch (FormatException fe)
                    {
                        Console.WriteLine($"Error parsing line: {line}. Details: {fe.Message}");
                    }
                    catch (IndexOutOfRangeException iore)
                    {
                        Console.WriteLine($"Error: Malformed line (too few columns): {line}. Details: {iore.Message}");
                    }
                }
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"Error reading CSV file {csvFilePath}: {ioEx.Message}");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred while processing CSV: {ex.Message}");
                return;
            }

            if (normalizedData.Any())
            {
                try
                {
                    using (var context = new UPMContext())
                    {
                        var partNumbers = normalizedData.Select(d => d.PartNumberName).Distinct().ToList();
                        var minDate = normalizedData.Min(d => d.ProductionDate).Date;
                        var maxDate = normalizedData.Max(d => d.ProductionDate).Date;

                        var existingRecords = await context.ProductionAchievements
                            .Where(p => partNumbers.Contains(p.PartNumberName) && p.ProductionDate >= minDate && p.ProductionDate <= maxDate)
                            .ToDictionaryAsync(p => $"{p.PartNumberName}|{p.ProductionDate:yyyy-MM-dd}");

                        int newRecordsCount = 0;
                        int updatedRecordsCount = 0;

                        foreach (var achievement in normalizedData)
                        {
                            var key = $"{achievement.PartNumberName}|{achievement.ProductionDate:yyyy-MM-dd}";

                            if (existingRecords.TryGetValue(key, out var existingRecord))
                            {
                                bool hasChanged = existingRecord.WorkingTime != achievement.WorkingTime ||
                                                  existingRecord.ProductionObjetive != achievement.ProductionObjetive ||
                                                  existingRecord.ProductionReal != achievement.ProductionReal ||
                                                  existingRecord.Supervisor != achievement.Supervisor ||
                                                  existingRecord.Leader != achievement.Leader ||
                                                  existingRecord.Shift != achievement.Shift;

                                if (hasChanged)
                                {
                                    existingRecord.WorkingTime = achievement.WorkingTime;
                                    existingRecord.ProductionObjetive = achievement.ProductionObjetive;
                                    existingRecord.ProductionReal = achievement.ProductionReal;
                                    existingRecord.Supervisor = achievement.Supervisor;
                                    existingRecord.Leader = achievement.Leader;
                                    existingRecord.Shift = achievement.Shift;
                                    existingRecord.CreateDate = DateTime.UtcNow;
                                    existingRecord.CreateBy = "System_Upsert";
                                    updatedRecordsCount++;
                                }
                            }
                            else
                            {
                                context.ProductionAchievements.Add(achievement);
                                newRecordsCount++;
                            }
                        }

                        await context.SaveChangesAsync();
                        Console.WriteLine($"Successfully loaded data. New records: {newRecordsCount}, Updated records: {updatedRecordsCount}.");
                    }
                }
                catch (DbUpdateException dbEx)
                {
                    Console.WriteLine($"Database update error: {dbEx.Message}");
                    if (dbEx.InnerException != null)
                    {
                        Console.WriteLine($"Inner exception: {dbEx.InnerException.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred while saving data to the database: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No valid production achievement records found to load into the database.");
            }
        }
    }
}


