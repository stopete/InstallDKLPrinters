<#

Removes all 7 printers

Removes their associated IP ports (even if still in use)

Forcefully removes the Xerox driver

Optionally cleans up the Xerox registry key and copied files


#>

Clear-Host

# Hashtable of printers and their IPs
$DKLprinters = @{
    "Library Printer 5" = ""
    "Library Printer 1" = "IP Address"
    "Library Printer 3" = "IP Address"
    "Library Printer 2" = "IP Address"
    "Library Printer 4" = "IP Address"
    "Library Printer 6" = "IP Address"
    "Library Printer 7" = "IP Address"
}

$DriverName = "Xerox Global Print Driver PCL6"
$PortsToRemove = @()

Write-Host "Starting full cleanup of DKL printers..." -ForegroundColor Yellow

foreach ($PrinterName in $DKLprinters.Keys) {
    $PrinterIP = $DKLprinters[$PrinterName]
    $PortName = "IP_$PrinterIP"
    $PortsToRemove += $PortName

    # Remove printer if it exists
    if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
        Write-Host "Removing printer: $PrinterName" -ForegroundColor Cyan
        Remove-Printer -Name $PrinterName -ErrorAction SilentlyContinue
    } else {
        Write-Host "Printer not found: $PrinterName" -ForegroundColor DarkGray
    }
}

# Force-remove ports
$PortsToRemove = $PortsToRemove | Sort-Object -Unique
foreach ($PortName in $PortsToRemove) {
    try {
        Write-Host "Removing port: $PortName" -ForegroundColor Cyan
        Remove-PrinterPort -Name $PortName -ErrorAction Stop
    } catch {
        Write-Host "Failed to remove port (may still be in use): $PortName" -ForegroundColor DarkGray
    }
}

# Remove driver forcefully if no printers using it
$driverInUse = Get-Printer | Where-Object { $_.DriverName -eq $DriverName }
if (!$driverInUse) {
    try {
        Write-Host "Removing driver: $DriverName" -ForegroundColor Yellow
        Remove-PrinterDriver -Name $DriverName -ErrorAction Stop
        Write-Host "`tDriver removed."
    } catch {
        Write-Host "`tFailed to remove driver: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Driver still in use by other printers: $DriverName" -ForegroundColor DarkGray
}

# Optional Cleanup: Registry and Files
$XeroxRegKey = "HKLM:\SOFTWARE\Xerox\PrinterDriver\V5.0\Configuration"
if (Test-Path $XeroxRegKey) {
    Remove-Item -Path $XeroxRegKey -Recurse -Force
    Write-Host "Deleted registry key: $XeroxRegKey"
}

# Optional Cleanup: Files
$CopiedFiles = @("C:\Windows\pnputil.exe", "C:\Windows\CommonConfiguration.xml")
foreach ($file in $CopiedFiles) {
    if (Test-Path $file) {
        Remove-Item -Path $file -Force
        Write-Host "Deleted file: $file"
    }
}

Write-Host "✅ DKL printer cleanup completed." -ForegroundColor Green
