$head = @"
 <meta http-equiv="refresh"
   content="60">
<link href="https://fonts.googleapis.com/css?family=Roboto" rel="stylesheet">

 <style>
html {
  margin: 0;
  height: 100%;
}
body {
  overflow:hidden;
  font-family: 'Roboto';
  font-size: 18pt;
  color:#FFF;
  background-color: #FF7F32;
  backgrond-image: url("sw_logo_wide.png");
  background-repat: none;
  }
table {
    text-align: left;
    width:98%; 
    margin-left:1%; 
    margin-right:1%;
    margin-top:-10px;
    border-collapse:collapse;
    border:0px;
    font-size: 14pt;
    
}
#footer {
    font-size: 8pt;
    color: rgba(255,255,255,0.5);
    margin-top: 10px;
    text-align: right;
}
</style>

"@
#endregion
#region page-header
#this is inserted after <body> and before <table>
#not used currently; not needed for aesthetics.
$body = @"
<div id="toplayer"></div>
<div id="faux-terminal">
  <div class="layer"></div>
  <div class="overlay"></div>
</div>
<div id="title"></div>
<div id="byline"><center><img src="sw_logo_wide.png" height=30% width=30%></center></div>
"@

#endregion

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


If ([IntPtr]::Size -eq 8) {Add-Type -Path "C:\Program Files\System.Data.SQLite\2010\bin\System.Data.SQLite.dll"}
If ([IntPtr]::Size -eq 4) {Add-Type -Path "C:\Program Files (x86)\System.Data.SQLite\2010\bin\System.Data.SQLite.dll"}

$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=C:\Program Files (x86)\Spiceworks\db\spiceworks_prod.db"
$con.Open()

$sql = $con.CreateCommand()
$sql.CommandText = "SELECT tickets.id AS ticket, UPPER(substr(users.email,1,2)) || substr(users.email,3,LENGTH(users.email)-18) AS creator, substr(Summary,1,60) AS Summary, tickets.c_status AS priority, ifnull(REPLACE(REPLACE(REPLACE(assigned_to, '1','Me'),'98','OtherTech'),'88','On Hold'),'N/A') AS assigned, substr(tickets.created_at, 6, 11) AS created_at, substr(tickets.updated_at,6,11) AS updated_at FROM tickets CROSS JOIN users WHERE tickets.created_by=users.id AND tickets.status<>'closed'"
$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
$data = New-Object System.Data.DataSet
[void]$adapter.Fill($data)

$con.Close()

$Footer = "<div id='footer'>Refresh completed at $Time</div>"

If ($context -eq "local") {
#$Outfile = "C:\Program Files (x86)\PRTG Network Monitor\webroot\images\sw.html"
}
Else {
#$Outfile = "\\opserver\c$\Program Files (x86)\PRTG Network Monitor\webroot\images\sw.html"
}

$data.Tables | Format-Table