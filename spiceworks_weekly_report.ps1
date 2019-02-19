#Pulls all open tickets, then all tickets closed within the week and sends them to a printer

#region Where am I running?
$Hostname = Get-Childitem env:computername
If ($Hostname.value -eq "opserver") {
$context = "local"
}
Else {
$context = "remote"
}
Write-Host "Script is running $($context)ly."
#endregion

Import-Module PSSQLite

Copy-Item "\\opserver\c$\Program Files (x86)\Spiceworks\db\spiceworks_prod.db" -Destination "C:\temp\sw-script\spiceworks_prod.db"
$Database = "C:\temp\sw-script\spiceworks_prod.db"
$Query = "SELECT tickets.id AS No, UPPER(substr(users.email,1,2)) || substr(users.email,3,LENGTH(users.email)-18) AS Creator, substr(Summary,1,47) AS Summary, ifnull(REPLACE(REPLACE(REPLACE(assigned_to, '1','Me'),'98','OtherTech'),'88','On Hold'),'N/A') AS Technician, substr(tickets.created_at, 6, 11) AS Submitted FROM tickets CROSS JOIN users WHERE tickets.created_by=users.id AND tickets.status<>'closed'"


$OpenTickets = Invoke-SqliteQuery -Query $Query -DataSource $Database | Format-Table | Out-String

$Body =@"
      ___________ _____ _   _   _____ _____ _____  _   __ _____ _____ _____ 
     |  _  | ___ \  ___| \ | | |_   _|_   _/  __ \| | / /|  ___|_   _/  ___|
     | | | | |_/ / |__ |  \| |   | |   | | | /  \/| |/ / | |__   | | \  --. 
     | | | |  __/|  __|| .   |   | |   | | | |    |    \ |  __|  | |   --. \
     \ \_/ / |   | |___| |\  |   | |  _| |_| \__/\| |\  \| |___  | | /\__/ /
      \___/\_|   \____/\_| \_/   \_/  \___/ \____/\_| \_/\____/  \_/ \____/ 
                                                                       
                                                                       

"@

$Body += $OpenTickets

$Query = "SELECT tickets.id AS ticket, UPPER(substr(users.email,1,2)) || substr(users.email,3,LENGTH(users.email)-18) AS creator, substr(Summary,1,60) AS Summary, substr(tickets.closed_at,6,11) AS closed_at_date FROM tickets CROSS JOIN users WHERE tickets.created_by=users.id AND tickets.closed_at > datetime('now', '-7 days')"

$ClosedTickets = Invoke-SqliteQuery -Query $Query -DataSource $Database | Format-Table | Out-String

$Body +=@"
 _____  _     _____ _____ ___________   _____ _____ _____  _   __ _____ _____ _____ 
/  __ \| |   |  _  /  ___|  ___|  _  \ |_   _|_   _/  __ \| | / /|  ___|_   _/  ___|
| /  \/| |   | | | \  --.| |__ | | | |   | |   | | | /  \/| |/ / | |__   | | \  --. 
| |    | |   | | | | --. \  __|| | | |   | |   | | | |    |    \ |  __|  | |   --. \
| \__/\| |___\ \_/ /\__/ / |___| |/ /    | |  _| |_| \__/\| |\  \| |___  | | /\__/ /
 \____/\_____/\___/\____/\____/|___/     \_/  \___/ \____/\_| \_/\____/  \_/ \____/ 
                                                                                    
                                                                                    
"@

$Body += $ClosedTickets
$Body | Out-Printer -Name "\\printserver\HP-CLJ-3600n-1"
Write-Host $Body
