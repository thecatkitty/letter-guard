#Requires -version 2.0

Function Get-Free-Letter {
  $letters = "DEFGHIJKLMNOPQRSTUVWXYZ"
  
  $i = 0
  While($i -lt $letters.Length) {
    If(@(Get-Volume | Where DriveLetter -eq $letters[$i]).Count -eq 0) {
	  Return $letters[$i]
	}
	$i++
  }
}

Function Check-And-Reassign-Letters {
  ForEach($partition In @(Get-Partition | Where DriveLetter)) {
    $letter = ($partition.DriveLetter + ":\") | Get-ChildItem | Where Name -eq "LetterGuard.conf" | Get-Content
  
    If($letter -eq $null) {
      Write-Output ("  Dysk " + @($partition.DriveLetter) + ": nie jest skonfigurowany, pomijam.")
    } ElseIf($letter -eq $partition.DriveLetter) {
      Write-Output ("  Litera dysku " + @($partition.DriveLetter) + ": jest spójna z konfiguracj¹.")
    } Else {
      Write-Output ("  Litera dysku " + @($partition.DriveLetter) + ": nie jest spójna z konfiguracj¹!")
	  Write-Output ("    ¯¹dana litera '" + $letter + "', a obecnie jest '" + @($partition.DriveLetter) + "'.")
	
      $occupant = @(Get-Partition | Where DriveLetter -eq $letter)
	  If($occupant.Count -ne 0) {
	    $tmplet = @(Get-Free-Letter)
	    $curlet = $partition.DriveLetter
	  
        Write-EventLog –LogName Application –Source "Celones LetterGuard" –EntryType Warning –EventID 21 –Message ("Litera dysku " + @($partition.DriveLetter) + ": nie jest spójna z konfiguracj¹! ¯¹dana litera '" + $letter + "', a obecnie jest '" + @($partition.DriveLetter) + "'. Zamiana liter dysków " + $letter + ": i " + $curlet + ":...")
	    Write-Output ("    Zamiana liter dysków " + $letter + ": i " + $curlet + ":...")
	    $occupant | Set-Partition -NewDriveLetter $tmplet
	    $partition | Set-Partition -NewDriveLetter $letter
	    $occupant | Set-Partition -NewDriveLetter $curlet
	  } Else {
	    Write-EventLog –LogName Application –Source "Celones LetterGuard" –EntryType Warning –EventID 21 –Message ("Litera dysku " + @($partition.DriveLetter) + ": nie jest spójna z konfiguracj¹! ¯¹dana litera '" + $letter + "', a obecnie jest '" + @($partition.DriveLetter) + "'. Przypisanie litery dysku " + $letter + ":...")
	    Write-Output ("    Przypisanie litery dysku " + $letter + ":...")
	    $partition | Set-Partition -NewDriveLetter $letter
	  }
	}
  }
}

Write-EventLog –LogName Application –Source "Celones LetterGuard" –EntryType Information –EventID 10 –Message "Uruchomiono us³ugê, rozpoczêto procedurê."
Write-Output ((Get-Date).ToString() + " Rozpoczêcie pracy...")
Check-And-Reassign-Letters
Register-WmiEvent -Class win32_VolumeChangeEvent -SourceIdentifier volumeChange
Do {
  $newEvent = Wait-Event -SourceIdentifier volumeChange
  $eventType = $newEvent.SourceEventArgs.NewEvent.EventType
  $eventTypeName = Switch($eventType) {
    1 {"zmiana konfiguracji"}
    2 {"dodanie urz¹dzenia"}
    3 {"usuniêcie urz¹dzenia"}
    4 {"dokowanie"}
  }
  
  Write-Output ((Get-Date).ToString() + " Wykry³em zdarzenie: " + $eventTypeName)
  If($eventType -eq 2) {
    Start-Sleep -Seconds 3
	Write-EventLog –LogName Application –Source "Celones LetterGuard" –EntryType Information –EventID 11 –Message "Dodano noœnik, rozpoczêto procedurê."
	Check-And-Reassign-Letters
  }
  Remove-Event -SourceIdentifier volumeChange
} While(1 -eq 1)
Unregister-Event -SourceIdentifier volumeChange
Write-EventLog –LogName Application –Source "Celones LetterGuard" –EntryType Information –EventID 15 –Message "Zakoñczono us³ugê."
Write-Output ((Get-Date).ToString() + " Zakoñczono...")
