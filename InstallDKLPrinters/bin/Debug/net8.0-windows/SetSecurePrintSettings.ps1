
# Registry key to check to make sure exist
$Key = "HKLM:\SOFTWARE\Xerox\PrinterDriver\V5.0\Configuration"


# make sure the key exists:
$exists = Test-Path -Path $key

if (!$exists) { $null = New-Item -Path $key -Force

    # Create Registry values
    Set-ItemProperty HKLM:\SOFTWARE\Xerox\PrinterDriver\V5.0\Configuration -Name RepositoryUNCPath -Value "C:\Windows" -Type String
    
 }

 
 
 
 # Set the following reg key so the computer can get metadata once the 
# the printers are installed.
$RegKey ="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata\"
Set-ItemProperty -path $RegKey -name PreventDeviceMetadataFromNetwork -value 1
 
# Gets the working directory when running the script as a  powershell script
 
$workingdirectory = [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\') 
if ($workingdirectory -eq $PSHOME.TrimEnd('\')) 
{     
	$workingdirectory = $PSScriptRoot 

    
}
 
 
 $path = $workingdirectory + "\"
 
 
 # Set the following reg key so the computer can get metadata once the 
# the printers are installed.
$RegKey ="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata\"
Set-ItemProperty -path $RegKey -name PreventDeviceMetadataFromNetwork -value 1


$configxmlpath = $path + "CommonConfiguration.xml"

Copy-Item –Path $configxmlpath –Destination 'C:\Windows\' -Force -Confirm:$False











#CreateRegkey