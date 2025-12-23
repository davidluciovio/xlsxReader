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
    public class Stamp_OperationalEfficiency_LoadData
    {
        public static void NormalizeEstampado(string inputFile, string outputFile)
        {
            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            var normalizedData = new List<OperationalEfficiency>();

            // Variables de contexto que se irán actualizando al encontrar nuevos encabezados
            string[] currentSupervisors = null;
            string[] currentLeaders = null;
            string[] currentDates = null;

            for (int i = 0; i < lines.Length; i++)
            {
                var columns = csvSplitRegex.Split(lines[i]).Select(s => s.Trim(' ', '"')).ToArray();
                if (columns.Length < 5) continue;

                string indicator = columns[3]; // La columna 3 nos dice qué tipo de fila es

                // 1. Detectar y actualizar encabezados
                if (indicator.Equals("SUPERVISOR", StringComparison.OrdinalIgnoreCase))
                {
                    currentSupervisors = columns;
                    continue;
                }
                if (indicator.Equals("LIDER", StringComparison.OrdinalIgnoreCase))
                {
                    currentLeaders = columns;
                    continue;
                }
                // Detectar fila de fechas (usualmente empieza vacío en col 3 y tiene fechas en col 4)
                if (string.IsNullOrEmpty(indicator) && columns.Length > 4 && columns[4].Contains("/12/2025"))
                {
                    currentDates = columns;
                    continue;
                }

                // 2. Si encontramos "STROKE", comienza un bloque de datos de 9 filas
                if (indicator.Equals("STROKE", StringComparison.OrdinalIgnoreCase) && i + 8 < lines.Length)
                {
                    // Extraer las 9 filas del bloque
                    var groupLines = lines.Skip(i).Take(9).Select(l =>
                        csvSplitRegex.Split(l).Select(s => s.Trim(' ', '"')).ToArray()
                    ).ToList();

                    var area = groupLines[0][0];       // PRENSA (Col 0)
                    var shiftRaw = groupLines[0][1];   // TURNO (Col 1)
                    var partNumber = groupLines[0][2]; // NÚMERO DE PARTE (Col 2)

                    if (string.IsNullOrWhiteSpace(partNumber)) continue;

                    string shift = (shiftRaw == "1°") ? "1" : "3";

                    // Iterar por los días (Columnas 4 a 34)
                    for (int colIndex = 4; colIndex < (currentDates?.Length ?? 0); colIndex++)
                    {
                        var dateStr = currentDates[colIndex];
                        if (string.IsNullOrWhiteSpace(dateStr) || dateStr.Contains("APROV")) continue;

                        // Helper para obtener valores numéricos de las 9 filas
                        float GetVal(int rowIdx) => (colIndex < groupLines[rowIdx].Length && float.TryParse(groupLines[rowIdx][colIndex], NumberStyles.Any, CultureInfo.InvariantCulture, out float v)) ? v : 0;

                        float spmReal = GetVal(5); // Fila 6: SPM REAL
                        float spmSet = GetVal(6);  // Fila 7: SPM SET

                        if (spmReal > 0 || spmSet > 0)
                        {
                            var culture = CultureInfo.GetCultureInfo("es-MX");
                            var cleanDateStr = dateStr.Replace("a. m.", "AM").Replace("p. m.", "PM").Trim();

                            if (DateTime.TryParse(cleanDateStr, culture, DateTimeStyles.None, out DateTime prodDate))
                            {
                                normalizedData.Add(new OperationalEfficiency
                                {
                                    Id = Guid.NewGuid(),
                                    Active = true,
                                    CreateDate = DateTime.UtcNow,
                                    CreateBy = "System_Normalize_Estampado_AllShifts",
                                    ProductionDate = prodDate,
                                    Area = "ESTAMPADO",
                                    Supervisor = (currentSupervisors != null && colIndex < currentSupervisors.Length) ? currentSupervisors[colIndex] : "",
                                    Leader = (currentLeaders != null && colIndex < currentLeaders.Length) ? currentLeaders[colIndex] : "",
                                    Shift = shift,
                                    PartNumberName = partNumber,

                                    Hp = spmSet,
                                    PriductionReal = spmReal,
                                    TotalTime = spmSet,            // T.T.
                                    RealWorkingTime = spmReal,      // T.T.T.
                                    TotalDowntime = GetVal(2) + GetVal(3), // JUNTA + PILOTAJE
                                    OperativityPercent = spmReal / spmSet,   // % APROV.

                                    RealTime = GetVal(1),
                                    NoProgramabeDowntimeTime = GetVal(0) // STROKE
                                });
                            }
                        }
                    }

                    // Saltamos las 8 filas que ya procesamos en este bloque
                    i += 8;
                }
            }

            WriteEfficiencyToCsv(normalizedData, outputFile);
        }

        private static void WriteEfficiencyToCsv(List<OperationalEfficiency> data, string filePath)
        {
            var sb = new StringBuilder();
            // Header
            sb.AppendLine("Id,Active,CreateDate,CreateBy,ProductionDate,Area,Supervisor,Leader,Shift,PartNumberName,HP,Neck,RealTime,OperativityPercent,PriductionReal,TotalTime,ProgramabeDowntimeTime,RealWorkingTime,NetoWorkingTime,NetoProduictiveTime,TotalDowntime,NoProgramabeDowntimeTime,NoReportedTime,DowntimePercent,NoProgramableDowntimePercent,ProgramableDowntimePercent");

            foreach (var r in data)
            {
                var line = string.Format(CultureInfo.InvariantCulture,
                    "\"{0}\",\"{1}\",\"{2:s}\",\"{3}\",\"{4:yyyy-MM-dd}\",\"{5}\",\"{6}\",\"{7}\",\"{8}\",\"{9}\",{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}",
                    r.Id, r.Active, r.CreateDate, r.CreateBy, r.ProductionDate, r.Area, r.Supervisor, r.Leader, r.Shift, r.PartNumberName,
                    r.Hp, r.Neck, r.RealTime, r.OperativityPercent, r.PriductionReal, r.TotalTime, r.ProgramabeDowntimeTime,
                    r.RealWorkingTime, r.NetoWorkingTime, r.NetoProduictiveTime, r.TotalDowntime, r.NoProgramabeDowntimeTime,
                    r.NoReportedTime, r.DowntimePercent, r.NoProgramableDowntimePercent, r.ProgramableDowntimePercent);

                sb.AppendLine(line);
            }
            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

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

                    try
                    {
                        normalizedData.Add(new OperationalEfficiency
                        {
                            Id = Guid.NewGuid(), // Siempre nuevos IDs para recarga limpia
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.UtcNow,
                            CreateBy = "System_Reload_EF9",
                            ProductionDate = DateTime.Parse(columns[4], CultureInfo.InvariantCulture, DateTimeStyles.None),
                            Area = "ESTAMPADO",
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
    }
}
