$head = @"
 <meta http-equiv="refresh"
   content="60">
<link href='https://fonts.googleapis.com/css?family=VT323' rel='stylesheet' type='text/css'>
 <style>
 @font-face {
  font-family: 'VT323';
  font-style: normal;
  font-weight: normal;
  src: local('VT323'), url('VT323.woff') format('woff');
}
::selection {
  background: #0080FF;
  text-shadow: none !important;
}
html, body {
  margin: 0;
  height: 100%;
}
body {
  overflow:hidden;
  font-family: VT323, monospace;
  font-size: 18pt;
  color:#E7A336;
  
  background: -webkit-radial-gradient(center center, contain, rgba(116,50,31,1), rgba(73,25,4,1)) center center no-repeat, black; /*center center for the gradient scrolls with the page :(*/
  /*background: -webkit-radial-gradient(center 75%, contain, rgba(0,75,0,0.8), black) center center no-repeat, black;*/
  -webkit-background-size: 110% 100%;
}
table {
    text-align: left;
    width:90%; 
    margin-left:5%; 
    margin-right:5%;
    border-collapse:collapse;
    border:0px;

}

td {
white-space:nowrap;
}

th {
background-color: rgba(157,83,25,0.8);

}

tr:nth-child(odd){ background: rgba(244,221,85,0.2)}
#byline {
text-align:center;
font-size:54pt;
}
#title {
font-size:14pt;
}
#footer {
font-size:14pt;
}


#faux-terminal:before {
  // ... positioning
  z-index: 4010;
  background: linear-gradient(#444 50%, #000 50%);
  background-size: 100% 4px;
  background-repeat: repeat-y;
  opacity: .14;
  box-shadow : inset 0px 0px 1px 1px rgba(0, 0, 0, .8);
  animation: pulse 5s linear infinite;
}
 
@keyframes pulse {
  0%   {transform: scale(1.001);  opacity: .14; }
  8%   {transform: scale(1.000);  opacity: .13; }
  15%  {transform: scale(1.004);  opacity: .14; }
  30%  {transform: scale(1.002);  opacity: .11; }
  100% {transform: scale(1.000);  opacity: .14; }
}

#faux-terminal:after {
  // ... positioning
  z-index : 4011;
  background-color : $rudy-accent-color;
  background: radial-gradient(ellipse at center, rgba(0,0,0,1) 0%,rgba(0,0,0,0.62) 45%,rgba(0,9,4,0.6) 47%,$rudy-accent-color 100%);
  box-shadow : inset 0px 0px 4px 4px rgba(100, 100, 100, .5);
  opacity : .1;
}

.layer {
  // ... positioning
  z-index : 4001;
  box-shadow : inset 0px 0px 1px 1px rgba(64, 64, 64, .1);
  background: radial-gradient(ellipse at center,darken($rudy-accent-color,1%) 0%,rgba(64,64,64,0) 50%);
  transform-origin : 50% 50%;
  transform: perspective(20px) rotateX(.5deg) skewX(2deg) scale(1.03);
  animation: glitch 1s linear infinite;
  opacity: .9;
}
 
.layer:after {
  // ... positioning
  background: radial-gradient(ellipse at center, rgba(0,0,0,0.5) 0%,rgba(64,64,64,0) 100%);
  opacity: .1;
}
 
@keyframes glitch {
  0%   {transform: scale(1, 1.002); }
  50%   {transform: scale(1, 1.0001); }
  100% {transform: scale(1.001, 1); }
}



.overlay {
  // ... positioning
  z-index: 4100;
}
 
.overlay:before {
  content : '';
  position : absolute;
  top : 0px;
  width : 100%;
  height : 5px;
  background : #fff;
  background: linear-gradient(to bottom, rgba(255,0,0,0) 0%,rgba(255,250,250,1) 50%,rgba(255,255,255,0.98) 51%,rgba(255,0,0,0) 100%); /* W3C */
  opacity : .1;
  animation: vline 1.25s linear infinite;
}
 
.overlay:after {
  // ... positioning
  box-shadow: 0 2px 6px rgba(25,25,25,0.2),
              inset 0 1px rgba(50,50,50,0.1),
              inset 0 3px rgba(50,50,50,0.05),
              inset 0 3px 8px rgba(64,64,64,0.05),
              inset 0 -5px 10px rgba(25,25,25,0.1);
}
 
@keyframes vline {
  0%   { top: 0px;}
  100% { top: 100%;}
}

#toplayer {
position: absolute;
width:100%;
height:100%;
background-image:url("lines.png");
background-repeat: repeat;
z-index: 8000;
opacity: 0.2;
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
<div id="title">C:\HelpDesk\Tickets>display_tickets.exe -order "order-received" -status "not closed"</div>
<div id="byline">Current HelpDesk Tickets</div>
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
$sql.CommandText = "SELECT tickets.id AS ticket, UPPER(substr(users.email,1,2)) || substr(users.email,3,LENGTH(users.email)-18) AS creator, Summary, tickets.c_status AS priority, ifnull(REPLACE(REPLACE(REPLACE(assigned_to, '1','Me'),'98','OtherTech'),'88','On Hold'),'N/A') AS assigned, substr(tickets.created_at, 6, 11) AS created_at, substr(tickets.updated_at,6,11) AS updated_at FROM tickets CROSS JOIN users WHERE tickets.created_by=users.id AND tickets.status<>'closed'"
$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
$data = New-Object System.Data.DataSet
[void]$adapter.Fill($data)

$con.Close()

#replace ID# with names
$Time = Get-Date
$Footer = "<div id='footer'>display_tickets.exe completed at $Time </br>C:\HelpDesk\Tickets>"

If ($context -eq "local") {
$Outfile = "C:\inetpub\wwwroot\ops\spiceworks.html"
}
Else {
$Outfile = "\\opserver\c$\inetpub\wwwroot\ops\spiceworks.html"
}

##Return all of the rows and pipe it into the ConvertTo-HTML cmdlet, and then pipe that into our output file
$data.Tables | Select-Object -Expand Rows | Select -Property @{N='Number'; E={$_.ticket}}, @{N='Submitted By'; E={$_.creator}}, @{N='Ticket Summary'; E={$_.Summary}}, @{N='Technician'; E={$_.assigned}}, @{N='Date Submitted'; E={$_.created_at}}, @{N='Last Updated'; E={$_.updated_at}} |
ConvertTo-HTML -head $head -body $body -PostContent $Footer | Out-File $Outfile
