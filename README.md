

# 🖨️ Install Network Xerox Printers

A Windows Forms application that automates the installation of multiple network printers with appropriate drivers based on system architecture. Designed for DKL library environments with minimal user interaction.

---

## 🚀 Features

- Automatically detects system architecture (ARM64 or x64)
- Installs Xerox Global Print Driver PCL6
- Sets up printer ports and network printers via PowerShell
- Executes additional security settings using a PowerShell script
- Provides real-time progress updates via UI
- Includes a time/date/status strip for enhanced interface polish
- Prompts the user to reboot to finalize settings

---

## 🧰 Requirements

- **Windows OS** (tested on Windows 10/11)
- Administrator privileges
- `pnputil.exe` and `SetSecurePrintSettings.ps1` included in the working directory
- Includes 

---

## 📦 Included Printers

| Printer Name       | IP Address        |
|--------------------|------------------ |
| Library Printer 1  | Add IP Address    |
| Library Printer 2  | Add IP Address    |
| Library Printer 3  | Add IP Address    |
| Library Printer 4  | Add IP Address    |
| Library Printer 5  | Add IP Address    |
| Library Printer 6  | Add IP Address    |
| Library Printer 7  | Add IP Address    |

---

## 🛠️ How It Works

1. Press the **Install Printers** button.
2. The app:
   - Detects the architecture.
   - Copies required driver files.
   - Adds the Xerox driver via `pnputil`.
   - Creates printer ports and installs printers.
   - Runs the secure print settings script.
3. Prompts for a restart to apply changes.

---

## 🔐 Security Notes

- Uses PowerShell with `ExecutionPolicy Bypass` for executing the secure print script.
- Registry modification is disabled/commented but can be enabled for advanced configurations.

---

## 🧑‍💻 Developer Notes

- Code uses `RunCommand()` to execute shell commands and capture output/errors.
- Real-time progress is displayed on a `ProgressBar`.
- Architecture detection is handled via `System.Runtime.InteropServices`.

---

## 📅 UI Details

The app displays:
- Current date
- Current time
- A copyright

These are refreshed every second using `System.Windows.Forms.Timer`.

---

## ⚠️ Disclaimer

This tool is intended for use in trusted environments. Always verify printer driver files before deployment.
