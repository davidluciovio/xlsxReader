using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XSLXtoCSV.Service
{
    internal class CopyFile
    {
        public static void CopyFiles(string filePath, string destination)
        {
            try
            {
                var fileName = Path.GetFileName(filePath);
                var destinationFile = Path.Combine(destination, fileName);

                File.Copy(filePath, destinationFile, true);

                Console.WriteLine($"Archivo {filePath} copiado y reemplazado (si existía).");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR AL COPIAR EL ARCHIVO: '{filePath}', ERROR: '{ex.Message}'");
            }
        }
    }
}
