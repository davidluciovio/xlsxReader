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
    public class Assy01_OperationalEfficiency_LoadData
    {
        public static void NormalizeEfficiency(string inputFile, string outputFile, string area = "ENSAMBLE I", string shift = "1")
        {
            if (inputFile.Contains("MASTES.csv")) return;

            var lines = File.ReadAllLines(inputFile, Encoding.UTF8);

            if (lines.Length < 4)
            {
                throw new InvalidDataException("El archivo de eficiencia es demasiado corto.");
            }

            var csvSplitRegex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

            // Las fechas están en la fila 2 (índice 2) a partir de la columna 7
            var dateHeaderLine = lines[2];
            var dateHeaders = csvSplitRegex.Split(dateHeaderLine).Select(s => s.Trim(' ', '"')).ToArray();

            var normalizedData = new List<OperationalEfficiency>();

            // Los datos empiezan en la fila 3 (índice 3) y se procesan en bloques de 13 filas
            for (int i = 3; i + 12 < lines.Length; i += 13)
            {
                // Extraemos las 13 filas del bloque actual
                var groupLines = lines.Skip(i).Take(13).Select(l =>
                    csvSplitRegex.Split(l).Select(s => s.Trim(' ', '"')).ToArray()
                ).ToList();

                // Metadatos básicos (están en la primera fila del bloque)
                var supervisor = groupLines[0][1];
                var leader = groupLines[0][2];
                var partNumber = groupLines[0][3];

                // HP y Neck (están en la fila 2 y 3 del bloque, columna 4)
                float.TryParse(groupLines[1][4], NumberStyles.Any, CultureInfo.InvariantCulture, out float hp);
                float.TryParse(groupLines[2][4], NumberStyles.Any, CultureInfo.InvariantCulture, out float neck);

                if (string.IsNullOrWhiteSpace(partNumber)) continue;

                // Iterar por los días (Columnas 7 a 37 del CSV)
                for (int colIndex = 7; colIndex < dateHeaders.Length; colIndex++)
                {
                    var dateStr = dateHeaders[colIndex];
                    if (string.IsNullOrWhiteSpace(dateStr)) continue;

                    // Extraer valores de cada una de las 13 filas para el día actual
                    // Usamos un pequeño helper local para parsear floats
                    float GetVal(int rowIdx) => float.TryParse(groupLines[rowIdx][colIndex], NumberStyles.Any, CultureInfo.InvariantCulture, out float v) ? v : 0;

                    float realTime = GetVal(0); // TIEMPO REAL
                    float productionReal = GetVal(2); // PIEZAS

                    // Solo agregamos si hubo producción o tiempo reportado
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
                                CreateBy = "System_Normalize_Efficiency",
                                ProductionDate = prodDate.Month != DateTime.Now.Month ? shift == "1" ? new DateTime(prodDate.Year, prodDate.Day, prodDate.Month, 8,0,0) : new DateTime(prodDate.Year, prodDate.Day, prodDate.Month, 21, 30, 0) : shift == "1" ? prodDate.AddHours(8) : prodDate.AddHours(21.5),
                                Area = area,
                                Supervisor = supervisor,
                                Leader = leader,
                                Shift = shift,
                                PartNumberName = partNumber,
                                Hp = hp,
                                Neck = neck,

                                // Mapeo de indicadores según la estructura del archivo
                                RealTime = realTime,
                                OperativityPercent = GetVal(1),   // TAZA DE OPERATIVIDAD
                                PriductionReal = productionReal,  // PIEZAS
                                TotalTime = GetVal(3),            // TIEMPO TOTAL
                                ProgramabeDowntimeTime = GetVal(4), // PARO PROGRAMADO
                                RealWorkingTime = GetVal(5),      // TIEMPO REAL TRABAJADO
                                NetoWorkingTime = GetVal(6),      // TIEMPO NETO PRODUCTIVO (Mapeado a ambos)
                                NetoProduictiveTime = GetVal(6),  // TIEMPO NETO PRODUCTIVO
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
    }
}
