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
                    Console.WriteLine("CSV file is empty or contains only headers.");
                    return;
                }

                foreach (var line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var columns = csvSplitRegex.Split(line).Select(s => s.Trim(' ', '"')).ToArray();

                    if (float.Parse(columns[11], CultureInfo.InvariantCulture) == 0) continue;

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