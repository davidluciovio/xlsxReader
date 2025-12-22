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
    public class Assy02_OperationalEfficiency_LoadData
    {
        public static void NormalizeEfficiencyEnsII(string inputFile, string outputFile)
        {
            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);

            if (lines.Length < 4)
            {
                throw new InvalidDataException("El archivo de eficiencia ENS II es demasiado corto.");
            }

            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            // 1. Detectar Area y Turno desde la primera fila (Título)
            var titleRow = lines[0];
            string area = titleRow.Contains("ENSAMBLE II") ? "ENSAMBLE II" : "ENSAMBLE I";
            string shift = titleRow.Contains("1ER") ? "1" : (titleRow.Contains("3ER") ? "3" : "3");

            // 2. Las fechas están en la fila 1 (índice 1) a partir de la columna 6
            var dateHeaders = csvSplitRegex.Split(lines[1]).Select(s => s.Trim(' ', '"')).ToArray();

            var normalizedData = new List<OperationalEfficiency>();

            // 3. Los datos empiezan en la fila 3 (índice 3)
            // El archivo ENS II tiene bloques de 13 filas de datos + 1 fila vacía (total paso de 14)
            for (int i = 3; i + 12 < lines.Length; i += 14)
            {
                var groupLines = lines.Skip(i).Take(13).Select(l =>
                    csvSplitRegex.Split(l).Select(s => s.Trim(' ', '"')).ToArray()
                ).ToList();

                // Validar que la fila no sea solo comas o esté vacía
                if (groupLines[0].Length < 6 || string.IsNullOrWhiteSpace(groupLines[0][0])) continue;

                // Metadatos específicos de ENS II
                var partNumber = groupLines[0][0]; // El número de parte está en la primera columna
                var supervisor = groupLines[1][0]; // Supervisor en fila 2 del bloque
                var leader = groupLines[1][1];     // Líder en fila 2 del bloque

                // HP y Neck (Fila 2 y 3 del bloque, columna 3)
                float.TryParse(groupLines[1][3], NumberStyles.Any, CultureInfo.InvariantCulture, out float hp);
                float.TryParse(groupLines[2][3], NumberStyles.Any, CultureInfo.InvariantCulture, out float neck);

                // 4. Iterar por los días (Columnas 6 a 36 del CSV)
                for (int colIndex = 6; colIndex < dateHeaders.Length; colIndex++)
                {
                    var dateStr = dateHeaders[colIndex];
                    if (string.IsNullOrWhiteSpace(dateStr) || dateStr.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) continue;

                    // Helper para obtener valores numéricos de las filas del bloque
                    float GetVal(int rowIdx) => (colIndex < groupLines[rowIdx].Length && float.TryParse(groupLines[rowIdx][colIndex], NumberStyles.Any, CultureInfo.InvariantCulture, out float v)) ? v : 0;

                    float realTime = GetVal(0); // TIEMPO REAL
                    float productionReal = GetVal(2); // PIEZAS

                    // Solo procesar si hay datos ese día
                    if (realTime > 0 || productionReal > 0)
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
                                CreateBy = "System_Normalize_Efficiency_EnsII",
                                ProductionDate = prodDate,
                                Area = area,
                                Supervisor = supervisor,
                                Leader = leader,
                                Shift = shift,
                                PartNumberName = partNumber,
                                Hp = hp,
                                Neck = neck,

                                // Mapeo de KPIs (basado en el orden del archivo)
                                RealTime = realTime,
                                OperativityPercent = GetVal(1),   // TAZA DE OPERATIVIDAD
                                PriductionReal = productionReal,  // PIEZAS
                                TotalTime = GetVal(3),            // TIEMPO TOTAL
                                ProgramabeDowntimeTime = GetVal(4), // PARO PROGRAMADO
                                RealWorkingTime = GetVal(5),      // TIEMPO REAL TRABAJADO
                                NetoWorkingTime = GetVal(6),      // TIEMPO NETO PRODUCTIVO
                                NetoProduictiveTime = GetVal(6),
                                TotalDowntime = GetVal(7),        // TIEMPO DE PARO TOTAL
                                NoProgramabeDowntimeTime = GetVal(8), // PARO NO PROGRAMADO
                                NoReportedTime = GetVal(9),       // TIEMPO NO REPORTADO
                                DowntimePercent = GetVal(10),     // % DE PARO TOTAL
                                NoProgramableDowntimePercent = GetVal(11), // % DE PARO NO PROGRAMADO
                                ProgramableDowntimePercent = GetVal(12)    // % DE PARO NO REPORTADO
                            });
                        }
                    }
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

                    try
                    {
                        normalizedData.Add(new OperationalEfficiency
                        {
                            Id = Guid.Parse(columns[0]),
                            Active = bool.Parse(columns[1]),
                            CreateDate = DateTime.Parse(columns[2], CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal),
                            CreateBy = columns[3],
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
                try
                {
                    using (var context = new UPMContext())
                    {
                        // Definimos la clave única: Parte + Fecha + Turno
                        var partNumbers = normalizedData.Select(d => d.PartNumberName).Distinct().ToList();
                        var minDate = normalizedData.Min(d => d.ProductionDate).Date;
                        var maxDate = normalizedData.Max(d => d.ProductionDate).Date;

                        var existingRecords = await context.OperationalEfficiencies
                            .Where(p => partNumbers.Contains(p.PartNumberName) && p.ProductionDate >= minDate && p.ProductionDate <= maxDate)
                            .ToDictionaryAsync(p => $"{p.PartNumberName}|{p.ProductionDate:yyyy-MM-dd}|{p.Shift}");

                        int newRecords = 0;
                        int updatedRecords = 0;

                        foreach (var item in normalizedData)
                        {
                            var key = $"{item.PartNumberName}|{item.ProductionDate:yyyy-MM-dd}|{item.Shift}";

                            if (existingRecords.TryGetValue(key, out var existing))
                            {
                                // Verificar si hubo cambios en los KPIs principales
                                bool changed = existing.RealTime != item.RealTime ||
                                               existing.PriductionReal != item.PriductionReal ||
                                               existing.TotalDowntime != item.TotalDowntime ||
                                               existing.Supervisor != item.Supervisor;

                                if (changed)
                                {
                                    // Actualizar campos
                                    existing.Supervisor = item.Supervisor;
                                    existing.Leader = item.Leader;
                                    existing.Hp = item.Hp;
                                    existing.Neck = item.Neck;
                                    existing.RealTime = item.RealTime;
                                    existing.OperativityPercent = item.OperativityPercent;
                                    existing.PriductionReal = item.PriductionReal;
                                    existing.TotalTime = item.TotalTime;
                                    existing.ProgramabeDowntimeTime = item.ProgramabeDowntimeTime;
                                    existing.RealWorkingTime = item.RealWorkingTime;
                                    existing.NetoWorkingTime = item.NetoWorkingTime;
                                    existing.NetoProduictiveTime = item.NetoProduictiveTime;
                                    existing.TotalDowntime = item.TotalDowntime;
                                    existing.NoProgramabeDowntimeTime = item.NoProgramabeDowntimeTime;
                                    existing.NoReportedTime = item.NoReportedTime;
                                    existing.DowntimePercent = item.DowntimePercent;
                                    existing.NoProgramableDowntimePercent = item.NoProgramableDowntimePercent;
                                    existing.ProgramableDowntimePercent = item.ProgramableDowntimePercent;

                                    existing.CreateDate = DateTime.UtcNow;
                                    existing.CreateBy = "System_Upsert_Efficiency";
                                    updatedRecords++;
                                }
                            }
                            else
                            {
                                context.OperationalEfficiencies.Add(item);
                                newRecords++;
                            }
                        }

                        await context.SaveChangesAsync();
                        Console.WriteLine($"Carga finalizada. Nuevos: {newRecords}, Actualizados: {updatedRecords}.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error de base de datos: {ex.Message}");
                }
            }
        }
    }
}
