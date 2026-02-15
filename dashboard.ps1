# Import ticket data from csv file
$tickets = Import-Csv ".\tickets.csv"

#For each row, convert date columns into datetime format
$tickets | ForEach-Object {
    $_.CreatedDate = [datetime]$_.CreatedDate
    if ($_.ResolvedDate) {
        $_.ResolvedDate = [datetime]$_.ResolvedDate
    }
}

#Count total tickets, open tickets and closed tickets
$totalTickets = $tickets.Count
$openTickets = ($tickets | Where-Object {$_.Status -ne "Closed"}).Count
$closedTickets = ($tickets | Where-Object {$_.Status -eq "Closed"}).Count

# define SLA Breach counter variable
$breached = 0

#define total ticket resolution hours variable
$totalResolutionTime = 0

#For each ticket, if status = "Closed". Calculate the ticket's resolution time & add to total.
foreach ($ticket in $tickets) {
    if ($ticket.Status -eq "Closed") {
        $resolutionTime = ($ticket.ResolvedDate - $ticket.CreatedDate).TotalHours
        $TotalResolutionTime += $resolutionTime
         }
}

#For each ticket, if it's resolution time is greater than it's SLA hours, add +1 to the SLA Breach counter.
foreach ($ticket in $tickets) {
    if ($resolutionTime -gt $ticket.SLAHours) {
            $breached++
        }
}


#Calculate the average resolution time/hours of all tickets.
$avgResolutionTime = $totalResolutionTime / $closedTickets
#Round average resolution hours into 2 decimal places
$avgResolutionTime = [math]::Round($avgResolutionTime,2)


#Output results to screen.
Write-Output "Total Tickets: $totalTickets"
Write-Output "Open Tickets: $openTickets"
Write-Output "Closed Tickets: $closedTickets"
Write-Output "SLA Breaches: $breached"
Write-Output "Average Resolution Time: $avgResolutionTime hours"

#Generate HTML Dashboard
$html = @"
<html>
<head>
<title>Ticket Dashboard</title>
<style>
body { font-family: Arial; }
.card { padding: 15px; margin: 10px; border: 1px solid #ccc; display: inline-block; }
</style>
</head>
<body>
<h1>Service Desk Dashboard</h1>
<div class='card'>Total Tickets: $totalTickets</div>
<div class='card'>Open Tickets: $openTickets</div>
<div class='card'>Closed Tickets: $closedTickets</div>
<div class='card'>SLA Breaches: $breached</div>
<div class='card'>Avg Resolution Time: $avgResolutionTime hours</div>
</body>
</html>
"@
#Take the contents of $html and write it's output data into a file named "dashboard.html".
$html | Out-File ".\dashboard.html"

#Schedule the above script to be executed each Monday at 8am.
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
    -Argument "-File ./dashboard.ps1"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 8am

<#Creates the scheduled task in Windows Task Manager.
Register-ScheduledTask `
    -Action $action `
    -Trigger $trigger `
    -TaskName "Automated Excel Ticket Dashboard 2"
#>

#Email dashboard.html to currently logged in Outlook Account.
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Subject = "Weekly Ticket Dashboard"
$mail.HTMLBody = Get-Content "dashboard.html" -Raw
$mail.To = $outlook.Session.CurrentUser.Address



