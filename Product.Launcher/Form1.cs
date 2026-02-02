using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;


namespace Product.Launcher
{
    public partial class Form1 : Form
    {
        // MVP paths
        private const string WorkbookPath = @"C:\ProgramData\ProductMVP\Workbook\Workbook.xlsx";
        private const string LicensePath = @"C:\ProgramData\ProductMVP\License\license.txt";

        // Session token name (Excel Add-in will check this)
        private const string SessionEnvVar = "PRODUCT_SESSION_TOKEN";

        public Form1()
        {
            InitializeComponent();
        }

        private void btnLaunch_Click(object sender, EventArgs e)
        {
            try
            {
                // 1) Licence check (MVP)
                if (!File.Exists(LicensePath))
                {
                    MessageBox.Show("License file missing. Please contact admin.",
                        "License Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var licenseKey = File.ReadAllText(LicensePath).Trim();
                if (licenseKey != "VALID-KEY-123") // MVP key
                {
                    MessageBox.Show("Invalid license key.",
                        "License Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 2) Validate workbook exists
                if (!File.Exists(WorkbookPath))
                {
                    MessageBox.Show("Workbook not found:\n" + WorkbookPath,
                        "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3) Create session token (MVP)
                var token = Guid.NewGuid().ToString("N");

                // Store token in registry so Excel (and Excel-DNA addin) can read it
                Registry.SetValue(
                    @"HKEY_CURRENT_USER\Software\ProductMVP",
                    "SessionToken",
                    token,
                    RegistryValueKind.String
                );

                // Optional: also store expiry (UTC ticks)
                Registry.SetValue(
                    @"HKEY_CURRENT_USER\Software\ProductMVP",
                    "TokenExpiresUtcTicks",
                    DateTime.UtcNow.AddMinutes(2).Ticks.ToString(),
                    RegistryValueKind.String
                );


                // 4) Launch Excel and open workbook
                LaunchExcel(WorkbookPath);

                // Optional: show token for testing
                // MessageBox.Show("Token set: " + token);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void LaunchExcel(string workbookPath)
        {
            // Block if Excel already running (important)
            if (Process.GetProcessesByName("EXCEL").Length > 0)
            {
                MessageBox.Show("Please close all Excel windows before opening from Launcher.",
                    "Excel Running", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "excel.exe",
                Arguments = $"\"{workbookPath}\"",
                UseShellExecute = true
            });
        }

    }
}
