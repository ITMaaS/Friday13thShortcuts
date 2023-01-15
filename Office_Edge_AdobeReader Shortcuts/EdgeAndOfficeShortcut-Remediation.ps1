#Remediation script includes Office, Edge and Adobe Reader shortcuts.
#Credit to Call4Cloud.nl for publishing the original version of this script which I modified for my own needs
#Original version @ https://call4cloud.nl/wp-content/uploads/2023/01/repair.txt


#Restore Office Shortcuts to Public Start Menu

		if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Powerpoint.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\PowerPNT.exe"
		$ShortCut.Description = "PowerPoint"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Word.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Winword.EXE"
		$ShortCut.Description = "Word"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\Outlook.exe"
		$ShortCut.Description = "Outlook"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
		$ShortCut.Description = "Excel"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneNote.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\onenote.EXE"
		$ShortCut.Description = "OneNote"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}

	if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OneDrive.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft OneDrive\OneDrive.exe"
		$ShortCut.Description = "OneDrive"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}
		if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk")){  
	 $ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Access.lnk")
		$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\MSAccess.exe"
		$ShortCut.Description = "Access"
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
	}



#Restore Edge Shortcuts to 

if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")){  
 $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $ShortCut.Description = "Edge"
    $ShortCut.IconLocation = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe, 0"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.Save()

    $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\desktop\Microsoft Edge.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $ShortCut.Description = "Edge"
    $ShortCut.IconLocation = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe, 0"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.Save()
}

#Restore Other Shortcuts to Start Menu

    ###Adobe Reader
if(!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Acrobat Reader.lnk")){  
 $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Acrobat Reader.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
    $ShortCut.Description = "Acrobat Reader"
	$ShortCut.IconLocation = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe, 1"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.Save()
    }