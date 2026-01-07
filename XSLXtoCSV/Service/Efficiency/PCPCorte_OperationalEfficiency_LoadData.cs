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
    public class PCPCorte_OperationalEfficiency_LoadData
    {
        public static void NormalizeCorte(string inputFile, string outputFile)
        {
            // Lector robusto para manejar saltos de línea dentro de celdas
            string fullText = File.ReadAllText(inputFile, Encoding.UTF8);
            var rowSplitRegex = new Regex(@"\r?\n(?=(?:[^""]*""[^""]*"")*[^""]*$)");
            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            var lines = rowSplitRegex.Split(fullText).Where(l => !string.IsNullOrWhiteSpace(l)).ToArray();

            // Contexto dinámico
            string[] currentDates = null;
            string[] currentSupervisors = null;
            string[] currentLeaders = null;
            string[] currentShifts = null;

            int metadataColumnCount = 7; // Basado en el análisis de columnas: TIEMPO TACTO HP, ITEM, LINEA, etc.
            var normalizedData = new List<OperationalEfficiency>();

            for (int i = 0; i < lines.Length; i++)
            {
                var columns = csvSplitRegex.Split(lines[i]).Select(s => s.Trim(' ', '"', '\n', '\r')).ToArray();

                // 1. DETECCIÓN DE BLOQUE DE ENCABEZADOS
                if (columns.Length > 2 && columns[1] == "ITEM" && columns[2] == "LINEA")
                {
                    currentDates = columns; // Fila de Fechas

                    if (i + 1 < lines.Length)
                        currentSupervisors = csvSplitRegex.Split(lines[i + 1]).Select(s => s.Trim(' ', '"', '\n', '\r')).ToArray();

                    if (i + 2 < lines.Length)
                        currentLeaders = csvSplitRegex.Split(lines[i + 2]).Select(s => s.Trim(' ', '"').Replace("\n", " / ")).ToArray();

                    if (i + 4 < lines.Length)
                        currentShifts = csvSplitRegex.Split(lines[i + 4]).Select(s => {
                            var match = Regex.Match(s, @"\d+");
                            return match.Success ? match.Value : s;
                        }).ToArray();

                    i += 5; // Saltar el bloque de encabezados
                    continue;
                }

                // 2. PROCESAMIENTO DE DATOS
                // Validamos si la columna ITEM es numérica para identificar fila de datos
                if (currentDates != null && columns.Length > metadataColumnCount && int.TryParse(columns[1], out _))
                {
                    var partName = columns[2]; // Columna LINEA

                    // Metadatos de la pieza
                    float.TryParse(columns[3], NumberStyles.Any, CultureInfo.InvariantCulture, out float hp);
                    float.TryParse(columns[4], NumberStyles.Any, CultureInfo.InvariantCulture, out float targetTime);

                    // Iterar por bloques de 4 columnas (2 turnos por día)
                    for (int dayIdx = 0; dayIdx < (currentDates.Length - metadataColumnCount) / 4; dayIdx++)
                    {
                        for (int subIdx = 0; subIdx < 2; subIdx++) // 0 = 1er Turno, 1 = 3er Turno
                        {
                            int colIdx = metadataColumnCount + (dayIdx * 4) + (subIdx * 2);
                            if (colIdx + 1 >= columns.Length) break;

                            // T.T. REAL y % CUMP.
                            float.TryParse(columns[colIdx], NumberStyles.Any, CultureInfo.InvariantCulture, out float realVal);
                            float.TryParse(columns[colIdx + 1], NumberStyles.Any, CultureInfo.InvariantCulture, out float compliancePercent);

                            if (realVal > 0 || compliancePercent > 0)
                            {
                                string dateStr = currentDates[colIdx];
                                var culture = CultureInfo.GetCultureInfo("es-MX");
                                var cleanDate = dateStr.Replace("a. m.", "AM").Replace("p. m.", "PM").Trim();

                                if (DateTime.TryParse(cleanDate, culture, DateTimeStyles.None, out DateTime prodDate))
                                {
                                    normalizedData.Add(new OperationalEfficiency
                                    {
                                        Id = Guid.NewGuid(),
                                        Active = true,
                                        CreateDate = DateTime.UtcNow,
                                        CreateBy = "System_Normalize_TactTime_Efficiency",
                                        ProductionDate = prodDate,
                                        Area = "PCP CORTE",
                                        Supervisor = (colIdx < currentSupervisors?.Length) ? currentSupervisors[colIdx] : "S/S",
                                        Leader = (colIdx < currentLeaders?.Length) ? currentLeaders[colIdx] : "S/L",
                                        Shift = (colIdx < currentShifts?.Length) ? currentShifts[colIdx] : (subIdx == 0 ? "1" : "3"),
                                        PartNumberName = partName,

                                        // Mapeo de KPIs
                                        Hp = hp,
                                        TotalTime = targetTime, // T.TACTO TARGET
                                        RealTime = realVal,     // T.T. REAL
                                        OperativityPercent = compliancePercent, // % CUMP.
                                        PriductionReal = realVal,
                                        RealWorkingTime = targetTime, // Asumimos tiempo planeado
                                        NetoWorkingTime = targetTime,

                                        // Mapeo a campos de Estampado para consistencia
                                        Tt = targetTime,
                                        SpmSet = hp,
                                        Aprov = compliancePercent
                                    });
                                }
                            }
                        }
                    }
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
