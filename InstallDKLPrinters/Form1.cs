using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace InstallDKLPrinters
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitializeStatusStrip();
            InitializeTimer();

        }
        private async void btnInstallPrinters_Click(object sender, EventArgs e)
        {
            btnInstallPrinters.Enabled = false;
            progressBar1.Value = 0;
            lblProgress.Text = "Starting...";

            await Task.Run(() => InstallPrinters());

            lblProgress.Text = "✅ Done!";
            btnInstallPrinters.Enabled = true;
        }

        private void InstallPrinters()
        {
            try
            {
                string workingDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\') + "\\";

                string pnputilPath = Path.Combine(workingDir, "pnputil.exe");
               // string configXmlPath = Path.Combine(workingDir, "CommonConfiguration.xml");

                int arch = GetProcessorArchitecture();
                string driverInf, driverFolder;

                if (arch == 12)
                {
                    driverFolder = "UNIV_5.1055.3.0_PCL6_ARM64_Driver.inf";
                    driverInf = Path.Combine(workingDir, driverFolder, "x3UNIVX.inf");
                    Log("Detected ARM64 - using ARM driver");
                }
                else
                {
                    driverFolder = "UNIV_5.759.5.0_PCL6_x64_Driver.inf";
                    driverInf = Path.Combine(workingDir, driverFolder, "x3UNIVX.inf");
                    Log("Detected x64 - using x64 driver");
                }

                

                string driverName = "Xerox Global Print Driver PCL6";

                File.Copy(pnputilPath, @"C:\Windows\pnputil.exe", true);
                //File.Copy(configXmlPath, @"C:\Windows\CommonConfiguration.xml", true);

               // Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata", "PreventDeviceMetadataFromNetwork", 1);
                //using (RegistryKey key = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Xerox\PrinterDriver\V5.0\Configuration"))
               // {
                   // key.SetValue("RepositoryUNCPath", @"C:\Windows");
               // }

                RunCommand("C:\\Windows\\pnputil.exe", $"/add-driver \"{driverInf}\"");
                System.Threading.Thread.Sleep(8000);

                var printers = new Dictionary<string, string>
                {
                    { "Library Printer 5", "Add IP Address" },
                    { "Library Printer 1", "Add IP Address" },
                    { "Library Printer 3", "Add IP Address" },
                    { "Library Printer 2", "Add IP Address" },
                    { "Library Printer 4", "Add IP Address" },
                    { "Library Printer 6", "Add IP Address" },
                    { "Library Printer 7", "Add IP Address" }
                };

                int totalPrinters = printers.Count;
                int currentStep = 0;

                Log($"Your computer architecture is {arch}");
                Log($"");

                foreach (var pair in printers)
                {
                    
                    string printerName = pair.Key;
                    string ip = pair.Value;
                    string portName = "IP_" + ip;
                    
                    Log($"Installing {printerName}...");

                    RunCommand("powershell", $"-Command \"if (-Not (Get-PrinterPort -Name '{portName}' -ErrorAction SilentlyContinue)) {{ Add-PrinterPort -Name '{portName}' -PrinterHostAddress '{ip}' }}\"");
                    RunCommand("powershell", $"-Command \"Add-PrinterDriver -Name '{driverName}' -ErrorAction SilentlyContinue\"");
                    RunCommand("powershell", $"-Command \"Add-Printer -Name '{printerName}' -PortName '{portName}' -DriverName '{driverName}' -ErrorAction SilentlyContinue\"");

                    currentStep++;
                    int percent = (int)((currentStep / (float)totalPrinters) * 100);
                    UpdateProgress(percent, printerName);
                }

                string scriptPath = Path.Combine(workingDir, "SetSecurePrintSettings.ps1");
                RunCommand("powershell", $"-ExecutionPolicy Bypass -File \"{scriptPath}\"");
                Log("✅ SeSecurePrintSetings.ps1 script executed successfully.");

                DialogResult result = MessageBox.Show(
                    "Installation complete. To apply all printer settings, the computer needs to restart.\nDo you want to restart now?",
                    "Restart Required",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question
                );

                if (result == DialogResult.OK)
                {
                    Log("🔁 User confirmed restart. Restarting system...");
                    Process.Start("shutdown", "/r /t 3"); // Restart in 3 seconds
                }
                else
                {
                    Log("⏳ Restart cancelled by user. Please restart manually later.");
                }



                Log("✅ Completed installing DKL printers.");
            }
            catch (Exception ex)
            {
                Log("❌ ERROR: " + ex.Message);
            }
        }

        private int GetProcessorArchitecture()
        {
            Architecture arch = RuntimeInformation.OSArchitecture;

            if (arch == Architecture.X86)
                return 0;
            else if (arch == Architecture.X64)
                return 9;
            else if (arch == Architecture.Arm || arch == Architecture.Arm64)
                return 5;
            else
                return -1;
        }

        private void RunCommand(string exe, string args)
        {
            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                }
            };
            proc.Start();
            string output = proc.StandardOutput.ReadToEnd();
            string error = proc.StandardError.ReadToEnd();
            proc.WaitForExit();

            if (!string.IsNullOrWhiteSpace(output)) Log(output);
            if (!string.IsNullOrWhiteSpace(error)) Log("⚠️ " + error);
        }

        private void Log(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => Log(message)));
                return;
            }
            textBox1.AppendText($"{DateTime.Now:T} - {message}{Environment.NewLine}");
        }

        private void UpdateProgress(int value, string currentPrinter)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateProgress(value, currentPrinter)));
                return;
            }

            progressBar1.Value = Math.Min(value, 100);
            lblProgress.Text = $"Installing: {currentPrinter} ({value}%)";
        }


        private void InitializeStatusStrip()
        {
            // Create StatusStrip
            StatusStrip statusStrip1 = new StatusStrip();

            // Create labels
            ToolStripStatusLabel dateLabel = new ToolStripStatusLabel();
            ToolStripStatusLabel timeLabel = new ToolStripStatusLabel();
            ToolStripStatusLabel copyrightLabel = new ToolStripStatusLabel();

            // Set initial text
            dateLabel.Name = "dateLabel";
            dateLabel.Text = DateTime.Now.ToString("MMMM dd, yyyy");

            timeLabel.Name = "timeLabel";
            timeLabel.Text = DateTime.Now.ToString("hh:mm:ss tt");

            copyrightLabel.Name = "copyrightLabel";
            copyrightLabel.Text = "Topete © 2025";

            // Add spacing between items
            ToolStripStatusLabel spacer1 = new ToolStripStatusLabel() { Spring = true };
            ToolStripStatusLabel spacer2 = new ToolStripStatusLabel() { Spring = true };

            // Add items to StatusStrip
            statusStrip1.Items.Add(dateLabel);
            statusStrip1.Items.Add(spacer1);
            statusStrip1.Items.Add(timeLabel);
            statusStrip1.Items.Add(spacer2);
            statusStrip1.Items.Add(copyrightLabel);

            // Add StatusStrip to the form
            this.Controls.Add(statusStrip1);
        }

        // ... other code ...

        private void InitializeTimer()
        {
            timer1 = new System.Windows.Forms.Timer
            {
                Interval = 1000 // 1 second
            };
            timer1.Tick += Timer_Tick;
            timer1.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // Update date and time labels
            foreach (Control control in this.Controls)
            {
                if (control is StatusStrip statusStrip)
                {
                    foreach (ToolStripItem item in statusStrip.Items)
                    {
                        if (item.Name == "dateLabel")
                            item.Text = DateTime.Now.ToString("MMMM dd, yyyy");
                        else if (item.Name == "timeLabel")
                            item.Text = DateTime.Now.ToString("hh:mm:ss tt");
                    }
                }
            }
        }

    }
}
