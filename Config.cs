using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NET8AutomatedReports
{
    public class EmailSettings
    {
        public string From { get; set; } = string.Empty;
        public string To { get; set; } = string.Empty;
        public string Cc { get; set; } = string.Empty;
        public string Bcc { get; set; } = string.Empty;
    }

    public class Config
    {
        public string SQLServer { get; set; } = string.Empty;
        public string SQLDBName { get; set; } = string.Empty;
        public string SQLUsername { get; set; } = string.Empty;
        public string SQLPassword { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;

        public string SMTPServer { get; set; } = string.Empty;
        public int SMTPPort { get; set; }
        public string SMTPUsername { get; set; } = string.Empty;
        public string SMTPPassword { get; set; } = string.Empty;
        public string SMTPEncryptionMode { get; set; } = "Auto";
        public bool LogEnabled { get; set; } = true;

        public EmailSettings Email { get; set; } = new();

        [JsonConverter(typeof(DateTimeConverter))]
        public DateTime? ReportDate { get; set; }

        public static Config Load(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"No se encontró el archivo de configuración en: {path}");

            var json = File.ReadAllText(path);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            return JsonSerializer.Deserialize<Config>(json, options)
                   ?? throw new InvalidDataException("El archivo de configuración no es válido.");
        }
    }

    public class DateTimeConverter : JsonConverter<DateTime?>
    {
        public override DateTime? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            var str = reader.GetString();
            if (DateTime.TryParse(str, out var result))
                return result;

            return null;
        }

        public override void Write(Utf8JsonWriter writer, DateTime? value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value?.ToString("o"));
        }
    }
}
