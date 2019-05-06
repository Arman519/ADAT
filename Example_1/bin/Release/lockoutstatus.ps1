$import= Import-csv "C:\Service Desk Ticketing Tool\ticket_data.csv"
$EmployeeNumber = $import.AffectedUser

$program = "C:\Service Desk Ticketing Tool\lockoutstatus.exe"
$programArgs = "-u:$EmployeeNumber@fnfglobal.local"
Invoke-Command -ScriptBlock { & $program $programArgs }
