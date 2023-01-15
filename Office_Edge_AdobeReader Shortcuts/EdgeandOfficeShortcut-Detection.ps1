#Check if Edge shortcut is missing
$shortcutName="Microsoft Edge"

$StartMenuDir = Join-Path -Path $env:ALLUSERSPROFILE -ChildPath "Microsoft\Windows\Start Menu\Programs"



if (Test-Path -Path $(Join-Path $StartMenuDir "$shortcutName.lnk")){

    Exit 0
    }
else
    {Exit 1
}

#Remediation script will be triggered on the basis that if you're missing Edge, you're probably missing Office and Adobe shortcuts.  Remediation script includes Office, Edge and Adobe Reader shortcuts.