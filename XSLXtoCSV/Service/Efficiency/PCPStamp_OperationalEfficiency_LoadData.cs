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

namespace XSLXtoCSV.Service.Efficiency
{
    public class PCPStamp_OperationalEfficiency_LoadData
    {
        public static void NormalizeEstampado(string inputFile, string outputFile, string press = "")
        {
            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            var normalizedData = new List<OperationalEfficiency>();

            string[] currentSupervisors = null;
            string[] currentLeaders = null;
            string[] currentDates = null;

            for (int i = 0; i < lines.Length; i++)
            {
                var columns = csvSplitRegex.Split(lines[i]).Select(s => s.Trim(' ', '"', '\n', '\r')).ToArray();
                if (columns.Length < 5) continue;

                string metricLabel = columns[3]; // Columna de métrica (STROKE, T.T., etc.)

                // 1. Detección de Encabezados (Contexto Dinámico)
                if (metricLabel.Equals("SUPERVISOR", StringComparison.OrdinalIgnoreCase))
                {
                    currentSupervisors = columns;
                    continue;
                }
                if (metricLabel.Equals("LIDER", StringComparison.OrdinalIgnoreCase))
                {
                    currentLeaders = columns;
                    continue;
                }
                // Fila de fechas (basado en el número de día 1, 2, 3...)
                if (string.IsNullOrEmpty(metricLabel) && columns.Length > 4 && int.TryParse(columns[4], out int day1) && day1 == 1)
                {
                    currentDates = columns;
                    continue;
                }

                // 2. Procesar bloque de 9 filas al encontrar "STROKE"
                if (metricLabel.Equals("STROKE", StringComparison.OrdinalIgnoreCase) && i + 8 < lines.Length)
                {
                    var block = lines.Skip(i).Take(9).Select(l =>
                        csvSplitRegex.Split(l).Select(s => s.Trim(' ', '"')).ToArray()
                    ).ToList();

                    var prensa = block[0][0];
                    var shiftRaw = block[0][1];
                    var partNumber = block[0][2];

                    if (string.IsNullOrWhiteSpace(partNumber)) continue;
                    string shift = Regex.Match(shiftRaw, @"\d+").Value;

                    for (int colIdx = 4; colIdx <= 34; colIdx++) // Días 1 al 31
                    {
                        if (colIdx >= (currentDates?.Length ?? 0)) break;

                        // Helper para parsear números
                        float GetVal(int rowIdx) => (colIdx < block[rowIdx].Length && float.TryParse(block[rowIdx][colIdx], NumberStyles.Any, CultureInfo.InvariantCulture, out float v)) ? v : 0;

                        float spmReal = GetVal(5);
                        float spmSet = GetVal(6);

                        if (partNumber == "1ER TURNO") continue;

                        if (spmReal > 0 || spmSet > 0)
                        {
                            int day = int.Parse(currentDates[colIdx]);
                            var prodDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, day);

                            normalizedData.Add(new OperationalEfficiency
                            {
                                Id = Guid.NewGuid(),
                                Active = true,
                                CreateDate = DateTime.UtcNow,
                                CreateBy = "System_Normalize_Estampado",
                                ProductionDate = prodDate.Month != DateTime.Now.Month ? shift == "1" ? new DateTime(prodDate.Year, prodDate.Day, prodDate.Month, 8, 0, 0) : new DateTime(prodDate.Year, prodDate.Day, prodDate.Month, 21, 30, 0) : shift == "1" ? prodDate.AddHours(8) : prodDate.AddHours(21.5),
                                Area = "PCP ESTAMPADO",
                                Supervisor = (currentSupervisors != null && colIdx < currentSupervisors.Length) ? currentSupervisors[colIdx] : "",
                                Leader = (currentLeaders != null && colIdx < currentLeaders.Length) ? currentLeaders[colIdx] : "",
                                Shift = shift,
                                PartNumberName = $"{partNumber} - {prensa}",

                                // Mapeo solicitado
                                Stroke = GetVal(0),
                                Tt = GetVal(1),
                                Junta = GetVal(2),
                                Pilotaje = GetVal(3),
                                Ttt = GetVal(4),
                                SpmReal = spmReal,
                                SpmSet = spmSet,
                                StSpmSet = GetVal(7),
                                Aprov = GetVal(8),

                                // Mapeo de compatibilidad
                                Hp = spmSet,
                                PriductionReal = spmReal,
                                TotalTime = GetVal(1),
                                RealWorkingTime = GetVal(4),
                                TotalDowntime = GetVal(2) + GetVal(3),
                                OperativityPercent = GetVal(8)
                            });
                        }
                    }
                    i += 8;
                }
            }
            WriteEfficiencyToCsv(normalizedData, outputFile);
        }

        private static void WriteEfficiencyToCsv(List<OperationalEfficiency> data, string filePath)
        {
            var sb = new StringBuilder();

            // 1. Header actualizado con los nuevos campos de Estampado
            sb.AppendLine("Id,Active,CreateDate,CreateBy,ProductionDate,Area,Supervisor,Leader,Shift,PartNumberName," +
                          "HP,Neck,RealTime,OperativityPercent,PriductionReal,TotalTime,ProgramabeDowntimeTime," +
                          "RealWorkingTime,NetoWorkingTime,NetoProduictiveTime,TotalDowntime,NoProgramabeDowntimeTime," +
                          "NoReportedTime,DowntimePercent,NoProgramableDowntimePercent,ProgramableDowntimePercent," +
                          "Stroke,TT,Junta,Pilotaje,TTT,SPM_Real,SPM_Set,ST_SPM_Set,Aprov");

            foreach (var r in data)
            {

                if ( string.IsNullOrEmpty(Sanitize(r.Leader)))
                {
                    continue;
                }
                // 2. Formato de línea actualizado (Índices 0 al 34)
                // Usamos InvariantCulture para asegurar que los decimales usen punto (.) y no coma (,)
                var line = string.Format(CultureInfo.InvariantCulture,
                    "\"{0}\",\"{1}\",\"{2:s}\",\"{3}\",\"{4:yyyy-MM-dd}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\"," +
                    "{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}," +
                    "{26},{27},{28},{29},{30},{31},{32},{33},{34}",
                    r.Id,
                    r.Active,
                    r.CreateDate,
                    r.CreateBy,
                    r.ProductionDate,
                    r.Area,
                    Sanitize(r.Supervisor),
                    Sanitize(r.Leader),
                    r.Shift,
                    Sanitize(r.PartNumberName),
                    r.Hp,
                    r.Neck,
                    r.RealTime,
                    r.OperativityPercent,
                    r.PriductionReal,
                    r.TotalTime,
                    r.ProgramabeDowntimeTime,
                    r.RealWorkingTime,
                    r.NetoWorkingTime,
                    r.NetoProduictiveTime,
                    r.TotalDowntime,
                    r.NoProgramabeDowntimeTime,
                    r.NoReportedTime,
                    r.DowntimePercent,
                    r.NoProgramableDowntimePercent,
                    r.ProgramableDowntimePercent,
                    // Nuevos campos de Estampado
                    r.Stroke,
                    r.Tt,
                    r.Junta,
                    r.Pilotaje,
                    r.Ttt,
                    r.SpmReal,
                    r.SpmSet,
                    r.StSpmSet,
                    r.Aprov);

                sb.AppendLine(line);
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        // Helper para evitar que comas o comillas en los nombres rompan el CSV
        private static string Sanitize(string field) => field?.Replace("\"", "\"\"") ?? "";

        public static async Task LoadEfficiencyToDatabase(string csvFilePath)
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

                    if (float.Parse(columns[14], CultureInfo.InvariantCulture) == 0) continue;

                    if (columns[6] == null )
                    {
                        continue;
                    }

                    try
                    {
                        normalizedData.Add(new OperationalEfficiency
                        {
                            Id = Guid.NewGuid(), // Siempre nuevos IDs para recarga limpia
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.UtcNow,
                            CreateBy = "System_Reload_EF9",
                            ProductionDate = DateTime.Parse(columns[4], CultureInfo.InvariantCulture, DateTimeStyles.None),
                            Area = "PCP ESTAMPADO",
                            Supervisor = columns[6],
                            Leader = columns[7],
                            Shift = columns[8],
                            PartNumberName = columns[9],
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
                            ProgramableDowntimePercent = float.Parse(columns[25], CultureInfo.InvariantCulture),
                            // Campos específicos de Estampado (respecto al header del CSV)
                            Stroke = float.Parse(columns[26], CultureInfo.InvariantCulture),
                            Tt = float.Parse(columns[27], CultureInfo.InvariantCulture),
                            Junta = float.Parse(columns[28], CultureInfo.InvariantCulture),
                            Pilotaje = float.Parse(columns[29], CultureInfo.InvariantCulture),
                            Ttt = float.Parse(columns[30], CultureInfo.InvariantCulture),
                            SpmReal = float.Parse(columns[31], CultureInfo.InvariantCulture),
                            SpmSet = float.Parse(columns[32], CultureInfo.InvariantCulture),
                            StSpmSet = float.Parse(columns[33], CultureInfo.InvariantCulture),
                            Aprov = float.Parse(columns[34], CultureInfo.InvariantCulture)

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
                    // Iniciamos una transacción para garantizar la atomicidad (borrado + inserción)
                    using var transaction = await context.Database.BeginTransactionAsync();

                    try
                    {
                        // 1. Identificamos qué Áreas y Periodos (Año/Mes) vienen en el CSV
                        // Esto nos permite limpiar exactamente lo que vamos a reponer
                        var targets = normalizedData
                            .GroupBy(x => new { x.Area, x.ProductionDate.Year, x.ProductionDate.Month })
                            .Select(g => g.Key);

                        foreach (var target in targets)
                        {
                            // 2. EF9: Ejecutar Delete directo en la base de datos
                            await context.OperationalEfficiencies
                                .Where(p => p.Area == target.Area
                                         && p.ProductionDate.Year == target.Year
                                         && p.ProductionDate.Month == target.Month)
                                .ExecuteDeleteAsync();

                            Console.WriteLine($"Limpieza completada: Área {target.Area} - Periodo {target.Month}/{target.Year}");
                        }

                        // 3. Inserción masiva de los nuevos datos
                        await context.OperationalEfficiencies.AddRangeAsync(normalizedData);
                        await context.SaveChangesAsync();

                        // 4. Confirmamos la transacción
                        await transaction.CommitAsync();
                        Console.WriteLine($"Carga finalizada con éxito. Registros insertados: {normalizedData.Count}");
                    }
                    catch (Exception ex)
                    {
                        // En caso de error, el Rollback restaura los datos borrados en el paso 2
                        await transaction.RollbackAsync();
                        Console.WriteLine($"Error crítico durante la carga. Se realizó Rollback: {ex.Message}");
                    }
                }
            }
        }

        public static async Task LoadToDatabase(string csvFilePath)
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

                    if (float.Parse(columns[14], CultureInfo.InvariantCulture) == 0) continue; // Saltar registros sin producción

                    try
                    {
                        normalizedData.Add(new OperationalEfficiency
                        {
                            Id = Guid.NewGuid(), // Siempre nuevos IDs para recarga limpia
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.UtcNow,
                            CreateBy = "System_Reload_EF9",
                            ProductionDate = DateTime.Parse(columns[4], CultureInfo.InvariantCulture, DateTimeStyles.None),
                            Area = columns[5],
                            Supervisor = columns[6],
                            Leader = columns[7],
                            Shift = columns[8],
                            PartNumberName = columns[9],
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
                            .GroupBy(x => new { x.Area, x.ProductionDate.Year, x.ProductionDate.Month })
                            .Select(g => g.Key);

                        foreach (var target in targets)
                        {
                            // 2. EF9: Eliminación directa y ultra rápida en la DB
                            await context.OperationalEfficiencies
                                .Where(p => p.Area == target.Area
                                         && p.ProductionDate.Year == target.Year
                                         && p.ProductionDate.Month == target.Month)
                                .ExecuteDeleteAsync();

                            Console.WriteLine($"Limpieza exitosa: {target.Area} ({target.Month}/{target.Year})");
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
