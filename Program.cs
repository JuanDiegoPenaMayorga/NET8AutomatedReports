using NET8AutomatedReports;
using Serilog;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace NET8AutomatedReports
{
    internal class Program
    {
        static void Main()
        {
            var stopwatchGlobal = Stopwatch.StartNew();

            string baseDir = AppContext.BaseDirectory;
            string configPath = Path.Combine(baseDir, "config.json");
            string mailHtmlPath = Path.Combine(baseDir, "mail.html");
            string subjectTemplatePath = Path.Combine(baseDir, "Subject.html");

            var stopwatch = Stopwatch.StartNew();
            var config = Config.Load(configPath);
            stopwatch.Stop();
            Log.Debug("Tiempo en cargar configuración: {Tiempo}ms", stopwatch.ElapsedMilliseconds);

            if (config.LogEnabled)
            {
                AllocConsole();

                Log.Logger = new LoggerConfiguration()
                    .WriteTo.Console()
                    .WriteTo.File("script.log", rollingInterval: RollingInterval.Day)
                    .MinimumLevel.Debug()
                    .CreateLogger();

                Log.Information("==== Inicio de ejecución ====");
            }
            else
            {
                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Fatal()
                    .CreateLogger();
            }

            try
            {
                var reportDate = config.ReportDate ?? DateTime.Now;
                Log.Information("Fecha de reporte: {Fecha}", reportDate);

                string outputPath = config.OutputPath;
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                    Log.Information("Directorio creado: {OutputPath}", outputPath);
                }

                string suffix = GetDaySuffix(reportDate.Day);

                stopwatch.Restart();
                string subjectTemplateRaw = File.Exists(subjectTemplatePath)
                    ? File.ReadAllText(subjectTemplatePath)
                    : "Customer incident report for {Month} {Day}{Suffix} {Year}";
                stopwatch.Stop();
                Log.Debug("Tiempo en leer plantilla de asunto: {Tiempo}ms", stopwatch.ElapsedMilliseconds);

                string subject = subjectTemplateRaw
                    .Replace("{{Month}}", reportDate.ToString("MMMM"))
                    .Replace("{{Day}}", reportDate.Day.ToString())
                    .Replace("{{Suffix}}", suffix)
                    .Replace("{{Year}}", reportDate.Year.ToString());

                string fileNameClean = Regex.Replace(subject, "<.*?>", string.Empty);
                foreach (var c in Path.GetInvalidFileNameChars())
                    fileNameClean = fileNameClean.Replace(c, '-');

                string outputFile = Path.Combine(outputPath, fileNameClean + ".xlsx");
                Log.Information("Archivo de salida: {File}", outputFile);

                stopwatch.Restart();
                var runner = new SqlRunner(config);
                var allTables = runner.RunAllQueries(reportDate);
                stopwatch.Stop();
                Log.Debug("Tiempo en ejecutar consultas SQL: {Tiempo}ms", stopwatch.ElapsedMilliseconds);

                stopwatch.Restart();
                ExcelExporter.Export(outputFile, allTables);
                stopwatch.Stop();
                Log.Information("Excel generado correctamente en: {File}", outputFile);
                Log.Debug("Tiempo en exportar Excel: {Tiempo}ms", stopwatch.ElapsedMilliseconds);

                try
                {
                    stopwatch.Restart();
                    Log.Information("Enviando correo...");
                    var mailer = new SmtpMailer(config);
                    mailer.SendReportEmail(subject, mailHtmlPath, outputFile);
                    stopwatch.Stop();
                    Log.Information("Correo enviado correctamente.");
                    Log.Debug("Tiempo en enviar correo: {Tiempo}ms", stopwatch.ElapsedMilliseconds);
                }
                catch (Exception emailEx)
                {
                    Log.Error("Error al enviar el correo: {Error}", emailEx);
                }
            }
            catch (Exception ex)
            {
                Log.Error("Error en la ejecución: {Error}", ex);
            }
            finally
            {
                stopwatchGlobal.Stop();
                Log.Information("==== Fin de ejecución ====");
                Log.Debug("Tiempo total de ejecución: {Tiempo}ms", stopwatchGlobal.ElapsedMilliseconds);
                Log.CloseAndFlush();
            }
        }

        static string GetDaySuffix(int day)
        {
            return day switch
            {
                1 or 21 or 31 => "st",
                2 or 22 => "nd",
                3 or 23 => "rd",
                _ => "th"
            };
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern bool AllocConsole();

    }
}
