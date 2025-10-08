using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace NET8AutomatedReports
{
    public class SqlRunner
    {
        private readonly string _connectionString;
        private readonly string _queriesFolder;

        public SqlRunner(Config config)
        {
            _connectionString = $"Server={config.SQLServer};Database={config.SQLDBName};User ID={config.SQLUsername};Password={config.SQLPassword};TrustServerCertificate=True;";
            _queriesFolder = Path.Combine(AppContext.BaseDirectory, "Queries");
        }

        public Dictionary<string, DataTable> RunAllQueries(DateTime reportDate)
        {
            var results = new Dictionary<string, DataTable>();
            string dateSql = reportDate.ToString("yyyy-MM-dd");

            var queryDir = Path.Combine(AppContext.BaseDirectory, "Queries");
            var queryFiles = Directory.GetFiles(queryDir, "*.txt")
                                      .OrderBy(f => f);

            foreach (var file in queryFiles)
            {
                string rawQuery = File.ReadAllText(file);
                string query = rawQuery.Replace("{ReportDate}", dateSql);

                string fileName = Path.GetFileNameWithoutExtension(file);
                string cleanName = Regex.Replace(fileName, @"^\d+\.?", "");

                var table = ExecuteQuery(query);
                results.Add(cleanName, table);
            }

            return results;
        }


        private DataTable ExecuteQuery(string query)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            var dt = new DataTable();
            using var cmd = new SqlCommand(query, conn);
            using var adapter = new SqlDataAdapter(cmd);
            adapter.Fill(dt);

            return dt;
        }
    }
}
