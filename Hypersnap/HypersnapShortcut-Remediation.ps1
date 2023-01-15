try{
[CmdletBinding()]
$ShortcutTargetPath = "c:\program files (x86)\Hypersnap 7\Hprsnap7.exe"

$ShortcutDisplayName = "HyperSnap 7"

$PinToStart= "true"

$IconFile= "c:\program files (x86)\Hypersnap 7\Hprsnap7.exe, 0"

$ShortcutArguments= $null

$WorkingDirectory= "C:\WINDOWS\SYSTEM32\CONFIG\SYSTEMPROFILE\PICTURES"


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


            $startMenuDir= Join-Path $env:ALLUSERSPROFILE "Microsoft\Windows\Start Menu\Programs\HyperSnap 7"



  $destinationPath = Join-Path -Path $startmenudir -ChildPath "$shortcutDisplayName.lnk"
    Add-Shortcut -DestinationPath $destinationPath -ShortcutTargetPath $ShortcutTargetPath -WorkingDirectory $WorkingDirectory
     Write-Output "Created shortcut to $ShortcutDisplayName"
    }
Catch
{Write-Output "Failed to create shortcut for $ShortcutDisplayName"
Exit 1}