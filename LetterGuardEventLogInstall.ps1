#Requires -version 2.0
New-EventLog –LogName Application –Source "Celones LetterGuard"
Write-EventLog –LogName Application –Source "Celones LetterGuard" –EntryType Information –EventID 1 –Message "Zainstalowano Dziennik zdarzeñ programu LetterGuard."
