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
        public static async Task DeleteMontlyData()
        {
            using (var context = new UPMContext())
            {
                var currentDate = DateTime.UtcNow;
                var firstDayOfMonth = new DateTime(currentDate.Year, currentDate.Month, 1);
                var lastDayOfPreviousMonth = firstDayOfMonth.AddDays(-1);
                var targetMonth = lastDayOfPreviousMonth.Month;
                var targetYear = lastDayOfPreviousMonth.Year;
                // Borrar datos de ProductionAchievement del mes objetivo
                await context.ProductionAchievements
                    .Where(p => p.ProductionDate.Year == targetYear && p.ProductionDate.Month == targetMonth)
                    .ExecuteDeleteAsync();
                // Borrar datos de OperationalEfficiency del mes objetivo
                await context.OperationalEfficiencies
                    .Where(o => o.ProductionDate.Year == targetYear && o.ProductionDate.Month == targetMonth)
                    .ExecuteDeleteAsync();
                Console.WriteLine($"Datos del mes {targetMonth}/{targetYear} eliminados correctamente.");
            }
        }
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
            // Regex compilado para mayor velocidad
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", RegexOptions.Compiled);

            try
            {
                // Usamos StreamReader para no saturar memoria RAM
                using (var reader = new StreamReader(csvFilePath, Encoding.UTF8))
                {
                    // Leer cabecera y descartarla
                    string line = await reader.ReadLineAsync();

                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        try
                        {
                            var columns = csvSplitRegex.Split(line).Select(s => s.Trim(' ', '"')).ToArray();

                            // Validación de seguridad de índices
                            if (columns.Length < 26) continue;

                            // 1. CORRECCIÓN SHIFT VACÍO:
                            // Usamos TryParse para evitar crash si viene vacío ("")
                            float shiftValue = 0;
                            if (!float.TryParse(columns[8], NumberStyles.Any, CultureInfo.InvariantCulture, out shiftValue))
                            {
                                // Si el turno no es un número válido o está vacío, decidimos si saltar o poner default.
                                // En tu lógica original saltabas si > 3. Si es 0 o vacío, ¿qué hacemos?
                                // Asumiremos que si está vacío no es válido para insertar.
                                continue;
                            }
                            if (shiftValue > 3) continue;

                            // 2. CORRECCIÓN OPERATIVIDAD:
                            if (!float.TryParse(columns[13], NumberStyles.Any, CultureInfo.InvariantCulture, out float operativity) || operativity == 0) continue;

                            //3. EXISTE PARTNUMBER
                            if(string.IsNullOrEmpty(columns[9])) continue;
                            // Filtros de texto
                            if (columns[9] == "3ER TURNO - 3ER TURNO" || columns[9] == "1ER TURNO - 1ER TURNO") continue;

                            normalizedData.Add(new OperationalEfficiency
                            {
                                Id = Guid.NewGuid(),
                                Active = bool.Parse(columns[1]),
                                CreateDate = DateTime.UtcNow,
                                CreateBy = "System_Reload_EF9",
                                ProductionDate = DateTime.Parse(columns[4], CultureInfo.InvariantCulture),
                                Area = columns[5].ToUpper(),
                                Supervisor = columns[6].ToUpper(),
                                Leader = columns[7].ToUpper(),
                                Shift = columns[8], // Guardamos el string original, o usa shiftValue.ToString()
                                PartNumberName = columns[9].ToUpper(),
                                Hp = float.Parse(columns[10], CultureInfo.InvariantCulture),
                                Neck = float.Parse(columns[11], CultureInfo.InvariantCulture),
                                RealTime = float.Parse(columns[12], CultureInfo.InvariantCulture),
                                OperativityPercent = operativity,
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
                            Console.WriteLine($"Error parseando línea: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error leyendo CSV: {ex.Message}");
                return;
            }

            if (normalizedData.Any())
            {
                using (var context = new UPMContext())
                using (var transaction = await context.Database.BeginTransactionAsync())
                {
                    try
                    {
                        // 3. CORRECCIÓN LÓGICA DE BORRADO:
                        // Agrupamos por TODOS los campos clave, incluyendo FECHA EXACTA (Día)
                        var targets = normalizedData
                            .GroupBy(x => new { x.Area, x.ProductionDate, x.Shift, x.PartNumberName })
                            .Select(g => g.Key)
                            .ToList();

                        foreach (var target in targets)
                        {
                            // Borramos SOLO el registro específico de ese día y turno
                            await context.OperationalEfficiencies
                                .Where(p => p.Area == target.Area
                                         && p.PartNumberName == target.PartNumberName
                                         && p.ProductionDate == target.ProductionDate // Match exacto de fecha (incluye día)
                                         && p.Shift == target.Shift)
                                .ExecuteDeleteAsync();
                        }

                        await context.OperationalEfficiencies.AddRangeAsync(normalizedData);
                        await context.SaveChangesAsync();
                        await transaction.CommitAsync();

                        Console.WriteLine($"Carga completada. Registros insertados: {normalizedData.Count}");
                    }
                    catch (Exception ex)
                    {
                        await transaction.RollbackAsync();
                        Console.WriteLine($"Error crítico (Rollback): {ex.Message}");
                    }
                }
            }
        }
    }
}
