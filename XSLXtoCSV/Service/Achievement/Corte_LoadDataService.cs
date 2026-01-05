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
            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);

            if (lines.Length < 5)
            {
                throw new InvalidDataException("El archivo de Tiempo Tacto es demasiado corto.");
            }

            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            // Metadata: 6 columnas iniciales antes de los días
            int metadataColumnCount = 6;

            // --- ENCABEZADOS HORIZONTALES ---
            // Fila 0: Fechas
            var dateHeaders = csvSplitRegex.Split(lines[0]).Skip(metadataColumnCount).ToArray();
            // Fila 1: Supervisores (Sergio Ramirez, etc.)
            var supervisorRow = csvSplitRegex.Split(lines[1]).Skip(metadataColumnCount).ToArray();
            // Fila 2: LIDERES (Gerardo Portillo, Jaime Valadez, etc.) <-- ESTA ES LA CORRECCIÓN
            var leaderRow = csvSplitRegex.Split(lines[2]).Skip(metadataColumnCount).ToArray();

            var normalizedData = new List<ProductionAchievement>();

            // Los datos numéricos de las líneas empiezan después de las filas de encabezado (Skip 10 para limpiar ruido)
            for (int i = 10; i < lines.Length; i++)
            {
                var row = lines[i];
                if (string.IsNullOrWhiteSpace(row) || row.Trim().Replace(",", "").Length == 0) continue;

                string[] columns = csvSplitRegex.Split(row).Select(s => s.Trim(' ', '"', '\n', '\r')).ToArray();

                // Validamos que la columna de LINEA (índice 2) tenga contenido
                if (columns.Length <= metadataColumnCount || string.IsNullOrWhiteSpace(columns[2])) continue;

                var lineName = columns[2]; // NOMBRE DE LA LINEA / PARTE

                // Iterar por bloques de 4 columnas por día (Target, Real, Diff, %)
                for (int j = 0; j < dateHeaders.Length / 4; j++)
                {
                    int dataIndexInHeaders = j * 4;
                    int dataIndexInRow = metadataColumnCount + (j * 4);

                    if (dataIndexInRow + 1 >= columns.Length) break;

                    var dateStr = dateHeaders[dataIndexInHeaders];
                    if (string.IsNullOrWhiteSpace(dateStr)) continue;

                    // Parseo de Target (Objetivo) y Real (Producción)
                    float.TryParse(columns[dataIndexInRow], NumberStyles.Any, CultureInfo.InvariantCulture, out float target);
                    float.TryParse(columns[dataIndexInRow + 1], NumberStyles.Any, CultureInfo.InvariantCulture, out float real);

                    if (target > 0 || real > 0)
                    {
                        try
                        {
                            var culture = CultureInfo.GetCultureInfo("es-MX");
                            var cleanDateStr = dateStr.Replace("a. m.", "AM").Replace("p. m.", "PM").Trim();

                            if (DateTime.TryParse(cleanDateStr, culture, DateTimeStyles.None, out DateTime productionDate))
                            {
                                normalizedData.Add(new ProductionAchievement
                                {
                                    Id = Guid.NewGuid(),
                                    Active = true,
                                    CreateDate = DateTime.UtcNow,
                                    CreateBy = "System_Normalize_TactTime",
                                    ProductionDate = productionDate,
                                    // Capturamos Supervisor y Líder desde sus respectivas filas de encabezado
                                    Supervisor = (dataIndexInHeaders < supervisorRow.Length) ? supervisorRow[dataIndexInHeaders] : "S/S",
                                    Leader = (dataIndexInHeaders < leaderRow.Length) ? leaderRow[dataIndexInHeaders] : "S/L",
                                    Shift = "1",
                                    PartNumberName = lineName,
                                    WorkingTime = target,
                                    ProductionObjetive = target,
                                    ProductionReal = real,
                                    Area = "CORTE Y ENSAMBLE"
                                });
                            }
                        }
                        catch (Exception ex)
                        {
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


