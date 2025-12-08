using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Data.Sqlite;
using System.Text.Json;
using System.Windows.Forms;

namespace KioskApp
{
    public partial class Form1 : Form
    {
        private string dbPath;
        private string csvPath;

        public Form1()
        {
            InitializeComponent();

string sharedDir = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments),
    "KioskApp");

Directory.CreateDirectory(sharedDir);

dbPath = Path.Combine(sharedDir, "survey.db");
csvPath = Path.Combine(sharedDir, "survey_export.csv");

            this.FormBorderStyle = FormBorderStyle.None;
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;

            InitializeDatabase();
            InitializeAsync();
        }

        /* ================================
           DATABASE INITIALIZATION
        =================================*/
        private void InitializeDatabase()
        {
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
                    name TEXT,
                    email TEXT,
                    timestamp TEXT DEFAULT CURRENT_TIMESTAMP
                );
            ";

            cmd.ExecuteNonQuery();
        }

        /* ================================
           WEBVIEW2 INITIALIZATION (FIXED)
        =================================*/
        private async void InitializeAsync()
        {
            string userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KioskApp_WebView2"
            );

            Directory.CreateDirectory(userDataFolder);

            var env = await CoreWebView2Environment.CreateAsync(
                browserExecutableFolder: null,
                userDataFolder: userDataFolder
            );

            await webView21.EnsureCoreWebView2Async(env);

            string rootPath = Path.Combine(Application.StartupPath, "wwwroot");

            webView21.CoreWebView2.SetVirtualHostNameToFolderMapping(
                "app.local",
                rootPath,
                CoreWebView2HostResourceAccessKind.Allow
            );

            webView21.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

            webView21.CoreWebView2.Navigate($"https://app.local/index.html?v={DateTime.Now.Ticks}");
        }

        /* ================================
           FIXED MESSAGE HANDLING
        =================================*/
        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            string raw = e.WebMessageAsJson;

            try
            {
                SurveyResult data = null;

                if (raw.StartsWith("{"))
                {
                    data = JsonSerializer.Deserialize<SurveyResult>(raw);
                }
                else if (raw.StartsWith("\""))
                {
                    string unwrapped = JsonSerializer.Deserialize<string>(raw);
                    data = JsonSerializer.Deserialize<SurveyResult>(unwrapped);
                }

                if (data != null)
                {
                    SaveToDatabase(data);
                }
                else
                {
                    MessageBox.Show("WebView2 sent an invalid message:\n" + raw);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON parse error:\n" + ex.Message + "\n\nRaw:\n" + raw);
            }
        }

        /* ================================
           SAVE TO DATABASE
        =================================*/
        private void SaveToDatabase(SurveyResult r)
        {
            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
                INSERT INTO responses (role, experience, brought, knownResearch, name, email)
                VALUES ($r, $x, $b, $k, $n, $e)
            ";

            cmd.Parameters.AddWithValue("$r", r.role ?? "");
            cmd.Parameters.AddWithValue("$x", r.experience ?? "");
            cmd.Parameters.AddWithValue("$b", r.brought ?? "");
            cmd.Parameters.AddWithValue("$k", r.knownResearch ?? "");
            cmd.Parameters.AddWithValue("$n", r.name ?? "");
            cmd.Parameters.AddWithValue("$e", r.email ?? "");

            cmd.ExecuteNonQuery();

            ExportToCsv();
        }

        public class SurveyResult
        {
            public string role { get; set; }
            public string experience { get; set; }
            public string brought { get; set; }
            public string knownResearch { get; set; }
            public string name { get; set; }
            public string email { get; set; }
        }

        /* ================================
           CSV EXPORT
        =================================*/
        private void ExportToCsv()
        {
            string tempPath = csvPath + ".tmp";

            using var conn = new SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT role, experience, brought, knownResearch, name, email FROM responses";

            using (var sw = new StreamWriter(tempPath, false))
            {
                sw.WriteLine("Role,Experience,What brought you to this booth?,Known OSU research centers and institutes,Name,Email");

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    sw.WriteLine(
                        $"{Escape(reader.GetString(0))}," +
                        $"{Escape(reader.GetString(1))}," +
                        $"{Escape(reader.GetString(2))}," +
                        $"{Escape(reader.GetString(3))}," +
                        $"{Escape(reader.GetString(4))}," +
                        $"{Escape(reader.GetString(5))}"
                    );
                }
            }

            if (File.Exists(csvPath))
                File.Delete(csvPath);

            File.Move(tempPath, csvPath);
        }

        private string Escape(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "\"\"";

            return "\"" + s.Replace("\"", "\"\"") + "\"";
        }
    }
}
