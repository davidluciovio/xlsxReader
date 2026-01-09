using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XSLXtoCSV.Data;
using XSLXtoCSV.Data.UPM_System;

namespace XSLXtoCSV.Service
{
    public class LoadToDatabase
    {
        public static async Task LoadAchievementToDatabase(string csvFilePath)
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

        public static async Task LoadOperativityToDatabase(string csvFilePath)
        {
            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine($"Error: CSV file not found at {csvFilePath}");
                return;
            }

            var normalizedData = new List<OperationalEfficiency>();
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

                    if (float.Parse(columns[13], CultureInfo.InvariantCulture) == 0) continue; // Saltar registros sin operatividad

                    if (float.Parse(columns[8], CultureInfo.InvariantCulture) > 3) continue; // Saltar registros sin operatividad

                    if (columns[9] == "3ER TURNO - 3ER TURNO" || columns[9] == "1ER TURNO - 1ER TURNO") continue; // Saltar registros sin operatividad

                    try
                    {
                        normalizedData.Add(new OperationalEfficiency
                        {
                            Id = Guid.NewGuid(), // Siempre nuevos IDs para recarga limpia
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.UtcNow,
                            CreateBy = "System_Reload_EF9",
                            ProductionDate = DateTime.Parse(columns[4], CultureInfo.InvariantCulture, DateTimeStyles.None),
                            Area = columns[5].ToUpper(),
                            Supervisor = columns[6].ToUpper(),
                            Leader = columns[7].ToUpper(),
                            Shift = columns[8],
                            PartNumberName = columns[9].ToUpper(),
                            Hp = float.Parse(columns[10], CultureInfo.InvariantCulture),
                            Neck = float.Parse(columns[11], CultureInfo.InvariantCulture),
                            RealTime = float.Parse(columns[12], CultureInfo.InvariantCulture),
                            OperativityPercent = float.Parse(columns[13], CultureInfo.InvariantCulture),
                            PriductionReal = float.Parse(columns[14], CultureInfo.InvariantCulture),
                            TotalTime = float.Parse(columns[15], CultureInfo.InvariantCulture),
                            ProgramabeDowntimeTime = float.Parse(columns[16], CultureInfo.InvariantCulture),
                            RealWorkingTime = float.Parse(columns[17], CultureInfo.InvariantCulture),
                            NetoWorkingTime = float.Parse(columns[18], CultureInfo.InvariantCulture),
                            NetoProduictiveTime = float.Parse(columns[19], CultureInfo.InvariantCulture),
                            TotalDowntime = float.Parse(columns[20], CultureInfo.InvariantCulture),
                            NoProgramabeDowntimeTime = float.Parse(columns[21], CultureInfo.InvariantCulture),
                            NoReportedTime = float.Parse(columns[22], CultureInfo.InvariantCulture),
                            DowntimePercent = float.Parse(columns[23], CultureInfo.InvariantCulture),
                            NoProgramableDowntimePercent = float.Parse(columns[24], CultureInfo.InvariantCulture),
                            ProgramableDowntimePercent = float.Parse(columns[25], CultureInfo.InvariantCulture)

                        });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error parsing efficiency line: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading CSV: {ex.Message}");
                return;
            }

            if (normalizedData.Any())
            {
                using (var context = new UPMContext())
                {
                    // Usamos una transacción para garantizar que no borremos datos si la inserción falla
                    using var transaction = await context.Database.BeginTransactionAsync();

                    try
                    {
                        // 1. Identificamos qué Áreas y Periodos (Mes/Año) vienen en el CSV
                        var targets = normalizedData
                            .GroupBy(x => new { x.Area, x.ProductionDate.Year, x.ProductionDate.Month, x.PartNumberName})
                            .Select(g => g.Key);

                        foreach (var target in targets)
                        {
                            // 2. EF9: Eliminación directa y ultra rápida en la DB
                            await context.OperationalEfficiencies
                                .Where(p => p.Area == target.Area
                                        && p.PartNumberName == target.PartNumberName
                                         && p.ProductionDate.Year == target.Year
                                         && p.ProductionDate.Month == target.Month)
                                .ExecuteDeleteAsync();

                            Console.WriteLine($"Limpieza exitosa: {target.Area} ({target.Month}/{target.Year}) {target.PartNumberName}");
                        }

                        // 3. Inserción de los nuevos datos normalizados
                        await context.OperationalEfficiencies.AddRangeAsync(normalizedData);
                        await context.SaveChangesAsync();

                        // 4. Confirmar transacción
                        await transaction.CommitAsync();
                        Console.WriteLine($"Carga finalizada exitosamente. Total insertados: {normalizedData.Count}");
                    }
                    catch (Exception ex)
                    {
                        // Revertimos la eliminación en caso de error
                        await transaction.RollbackAsync();
                        Console.WriteLine($"Error crítico durante la carga (Rollback ejecutado): {ex.Message}");
                    }
                }
            }
        }
    }
}
