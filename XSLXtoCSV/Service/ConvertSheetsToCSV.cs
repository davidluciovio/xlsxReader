using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XSLXtoCSV.Service
{
    public class ConvertSheetsToCSV
    {
        public ConvertSheetsToCSV()
        {
            
        }
        public static void ProcessExcelFixed(string excelFilePath)
        {
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Iteramos por cada pestaña (hoja)
                    do
                    {
                        string sheetName = reader.Name;
                        Console.WriteLine($"Procesando: {sheetName}...");

                        // 1. Crear un DataTable manualmente para la hoja actual
                        DataTable table = new DataTable();

                        // Cargamos los datos fila por fila
                        bool isFirstRow = true;
                        while (reader.Read())
                        {
                            // Crear columnas en la primera iteración
                            if (isFirstRow)
                            {
                                for (int i = 0; i < reader.FieldCount; i++)
                                    table.Columns.Add("Col" + i);
                                isFirstRow = false;
                            }

                            DataRow row = table.NewRow();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                row[i] = reader.GetValue(i);
                            }
                            table.Rows.Add(row);
                        }

                        // 2. Aplicar lógica de celdas combinadas ANTES de guardar
                        var mergedCells = reader.MergeCells;
                        if (mergedCells != null)
                        {
                            foreach (var range in mergedCells)
                            {
                                // Validar que el rango esté dentro de los datos leídos
                                if (range.FromRow < table.Rows.Count && range.FromColumn < table.Columns.Count)
                                {
                                    var valueToRepeat = table.Rows[range.FromRow][range.FromColumn];

                                    int maxRow = Math.Min(range.ToRow, table.Rows.Count - 1);
                                    int maxCol = Math.Min(range.ToColumn, table.Columns.Count - 1);

                                    for (int r = range.FromRow; r <= maxRow; r++)
                                    {
                                        for (int c = range.FromColumn; c <= maxCol; c++)
                                        {
                                            table.Rows[r][c] = valueToRepeat;
                                        }
                                    }
                                }
                            }
                        }

                        // 3. Guardar en CSV
                        string safeName = string.Join("_", sheetName.Split(Path.GetInvalidFileNameChars()));
                        string outputPath = $"{Path.GetFullPath(excelFilePath)}_{safeName}.csv";
                        SaveToCsv(table, outputPath);

                        // NextResult() ahora funcionará correctamente porque el cursor está al final de la hoja actual
                    } while (reader.NextResult());
                }
            }
        }
        static void SaveToCsv(DataTable table, string path)
        {
            var sb = new StringBuilder();
            foreach (DataRow row in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    string value = row[i]?.ToString() ?? "";
                    if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                        value = $"\"{value.Replace("\"", "\"\"")}\"";

                    sb.Append(value + (i < table.Columns.Count - 1 ? "," : ""));
                }
                sb.AppendLine();
            }
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
            Console.WriteLine($" -> {path} generado.");
        }
    }
}
