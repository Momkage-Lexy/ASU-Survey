using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Data.Sqlite;
using System.Text.Json;
using System.Windows.Forms;

namespace KioskApp
{
    public partial class Form1 : Form
    {
        // Writable paths for MSIX
        private string dbPath;
        private string csvPath;

        public Form1()
        {
            InitializeComponent();

            // ---------------------------------------------
            //  MSIX SAFE FILE STORAGE
            //  (Create writable folder under LocalAppData)
            // ---------------------------------------------
            string dataDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string kioskFolder = Path.Combine(dataDir, "KioskApp");
            Directory.CreateDirectory(kioskFolder);

            dbPath = Path.Combine(kioskFolder, "survey.db");
            csvPath = Path.Combine(kioskFolder, "survey_export.csv");

            InitializeDatabase();
            InitializeAsync();
        }

        // ============================================================
        // DATABASE INITIALIZATION
        // ============================================================
        private void InitializeDatabase()
        {
            // Create DB if needed
            if (!File.Exists(dbPath))
            {
                using var connection = new SqliteConnection($"Data Source={dbPath}");
                connection.Open();
            }

            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS responses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    role TEXT,
                    experience TEXT,
                    brought TEXT,
                    knownResearch TEXT,
                    email TEXT
                );
            ";

            cmd.ExecuteNonQuery();
        }

        // ============================================================
        // WEBVIEW2 INITIALIZATION
        // ============================================================
        private async void InitializeAsync()
        {
            await webView21.EnsureCoreWebView2Async(null);

            // (1) USE VIRTUAL HOST MAPPING — REQUIRED FOR MSIX
            string rootPath = Path.Combine(Application.StartupPath, "wwwroot");

            webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.local",
                rootPath,
                CoreWebView2HostResourceAccessKind.Allow
            );

            // (2) Load the offline HTML
            webView21.CoreWebView2.Navigate("https://app.local/index.html");

            // (3) Receive messages from the webpage
            webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;
        }

        // Handle messages from JS → C#
        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var json = e.WebMessageAsJson;
            var data = JsonSerializer.Deserialize<SurveyResult>(json);
            SaveToDatabase(data);
        }

        // ============================================================
        // DATABASE INSERT
        // ============================================================
        private void SaveToDatabase(SurveyResult r)
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO responses (role, experience, brought, knownResearch, email)
                VALUES ($r, $x, $b, $k, $e)
            ";

            cmd.Parameters.AddWithValue("$r", r.role);
            cmd.Parameters.AddWithValue("$x", r.experience);
            cmd.Parameters.AddWithValue("$b", r.brought);
            cmd.Parameters.AddWithValue("$k", r.knownResearch ?? "");
            cmd.Parameters.AddWithValue("$e", r.email ?? "");

            cmd.ExecuteNonQuery();

            ExportToCsv();
        }

        // Strongly typed model
        public class SurveyResult
        {
            public string role { get; set; }
            public string experience { get; set; }
            public string brought { get; set; }
            public string email { get; set; }
            public string knownResearch { get; set; }
        }

        // ============================================================
        // CSV EXPORT (MSIX SAFE)
        // ============================================================
        private void ExportToCsv()
        {
            string tempPath = csvPath + ".tmp";

            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT role, experience, brought, knownResearch, email FROM responses";

            using (var sw = new StreamWriter(tempPath, false))
            {
                sw.WriteLine("Role,Experience,What brought you to this booth?,Known OSU research centers and institutes,Email");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    sw.WriteLine(
                        $"{Escape(reader.GetString(0))}," +
                        $"{Escape(reader.GetString(1))}," +
                        $"{Escape(reader.GetString(2))}," +
                        $"{Escape(reader.GetString(3))}," +
                        $"{Escape(reader.GetString(4))}" 
                    );
                }
            }

            // Replace existing CSV safely
            if (File.Exists(csvPath))
                File.Delete(csvPath);

            File.Move(tempPath, csvPath);
        }

        // Escape CSV values
        private string Escape(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "\"\"";

            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }
    }
}
