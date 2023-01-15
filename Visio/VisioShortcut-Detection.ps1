##Run as System using Proactive Remediations


# Replace with the name of your shortcut (without *.lnk)
$shortcutName="Visio"
#Define Start Menu Location
$StartMenuDir = Join-Path -Path $env:ALLUSERSPROFILE -ChildPath "Microsoft\Windows\Start Menu\Programs"
#Define executable path
$ExectuablePath = "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"

#Check if exectuable exists
if (Test-Path -Path $ExectuablePath) {
#If executable exusts, check if shortcut is missing
if (Test-Path -Path $(Join-Path $StartMenuDir "$shortcutName.lnk")){
#Declare no action required if found
    Write-Host "Shortcut for $shortcutName found, no action required"
    Exit 0
    }
else
# Declare remediation required, exit 1 to flag for remediation in Proactive Remediations
    {Write-Output "Shortcut for $shortcutName not found, start remediation"
    Exit 1
}
}
#Declare that executable isn't found, so shortcut creation isn't required.
else {Write-Output "$shortcutName Executable Not Found, no action required"
Exit 0}