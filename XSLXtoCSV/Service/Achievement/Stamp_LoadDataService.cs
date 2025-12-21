using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XSLXtoCSV.Data.UPM_System;
using XSLXtoCSV.Data;
using Microsoft.EntityFrameworkCore;

namespace XSLXtoCSV.Service.Achievement
{
    public static class Stamp_LoadDataService
    {
        public static void NormalizeSPM(string inputFile, string outputFile, int year = 2025, int month = 12)
        {
            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);

            if (lines.Length < 4)
            {
                throw new InvalidDataException("El archivo SPM es demasiado corto.");
            }

            // Regex para CSV que respeta comillas
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            // Los datos comienzan en la fila 3 (índice 3)
            // Procesamos en bloques de 3 filas (PLAN, REAL, %)
            var normalizedData = new List<ProductionAchievement>();

            for (int i = 3; i < lines.Length; i += 3)
            {
                if (i + 1 >= lines.Length) break; // Necesitamos al menos PLAN y REAL

                var planRow = csvSplitRegex.Split(lines[i]).Select(s => s.Trim(' ', '"')).ToArray();
                var realRow = csvSplitRegex.Split(lines[i + 1]).Select(s => s.Trim(' ', '"')).ToArray();

                if (planRow.Length < 5) continue;

                // Metadatos comunes a las filas de este bloque
                var supervisor = planRow[0];
                var leader = planRow[1];
                var shiftRaw = planRow[2]; // Viene como "1ER", "3ER", etc.
                var prensa = planRow[3];   // Usaremos esto como PartNumberName

                // Limpiar el turno (Ej: "1ER" -> "1")
                var shift = Regex.Match(shiftRaw, @"\d+").Value;
                if (string.IsNullOrEmpty(shift)) shift = "1";

                // Las columnas de los días van de la 5 a la 35 (Días 1 al 31)
                for (int dayIndex = 5; dayIndex <= 35; dayIndex++)
                {
                    if (dayIndex >= planRow.Length || dayIndex >= realRow.Length) break;

                    var day = dayIndex - 4; // Columna 5 es Día 1
                    var objectiveStr = planRow[dayIndex];
                    var realStr = realRow[dayIndex];

                    bool hasObj = float.TryParse(objectiveStr, NumberStyles.Any, CultureInfo.InvariantCulture, out float objective);
                    bool hasReal = float.TryParse(realStr, NumberStyles.Any, CultureInfo.InvariantCulture, out float real);

                    // Solo agregamos si hay un valor real o un objetivo mayor a cero
                    if (objective > 0 || real > 0)
                    {
                        try
                        {
                            var productionDate = new DateTime(year, month, day);

                            normalizedData.Add(new ProductionAchievement
                            {
                                Id = Guid.NewGuid(),
                                Active = true,
                                CreateDate = DateTime.UtcNow,
                                CreateBy = "System_Normalize_SPM",
                                ProductionDate = productionDate,
                                Supervisor = supervisor,
                                Leader = leader,
                                Shift = shift,
                                PartNumberName = prensa,
                                WorkingTime = 0, // No disponible en este formato
                                ProductionObjetive = objective,
                                ProductionReal = real,
                                Area = "ESTAMPADO"
                            });
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            // Por si el mes tiene menos de 31 días y el CSV trae columnas extra
                            continue;
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