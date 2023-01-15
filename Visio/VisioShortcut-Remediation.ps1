try{
[CmdletBinding()]
$ShortcutTargetPath = "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"

$ShortcutDisplayName = "Visio"

$PinToStart= "true"

$IconFile= "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE, 0"

$ShortcutArguments= $null

$WorkingDirectory= $null


#helper function to avoid uneccessary code
function Add-Shortcut {
    param (
        [Parameter(Mandatory)]
        [String]$ShortcutTargetPath,
        [Parameter(Mandatory)]
        [String] $DestinationPath,
        [Parameter()]
        [String] $WorkingDirectory
    )

    process{
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($destinationPath)
        $Shortcut.TargetPath = $ShortcutTargetPath
        $Shortcut.Arguments = $ShortcutArguments
        $Shortcut.WorkingDirectory = $WorkingDirectory
    
        if ($IconFile){
            $Shortcut.IconLocation = $IconFile
        }
        # Create the shortcut
        $Shortcut.Save()
        #cleanup
        [Runtime.InteropServices.Marshal]::ReleaseComObject($WshShell) | Out-Null
    }
}


            $startMenuDir= Join-Path $env:ALLUSERSPROFILE "Microsoft\Windows\Start Menu\Programs"



  $destinationPath = Join-Path -Path $startmenudir -ChildPath "$shortcutDisplayName.lnk"
    Add-Shortcut -DestinationPath $destinationPath -ShortcutTargetPath $ShortcutTargetPath -WorkingDirectory $WorkingDirectory
    Write-Output "Created shortcut to $ShortcutDisplayName"
    Exit 0
}
Catch
{Write-Output "Failed to create shortcut for $ShortcutDisplayName"
Exit 1}