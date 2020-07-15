$recipient = "WSUS Master <wsus.master@example.com>"

$server = "wsus01.example.com"
$mailserver = "mail.example.com"
$mailfrom = "WSUS report $(HOSTNAME) <wsusreport@example.com>"

$bodycss = @"
<style>
h1, h2, h5, h6, th { text-align: center; font-family: Segoe UI; }
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #464646; color: #fff; max-width: 400px; padding: 5px 5px; }
td { font-size: 11px; padding: 5px 5px; color: #000; }
tr { background: #efefef; }
tr:nth-child(even) { background: #f3f3f3; }
tr:nth-child(odd) { background: #e4e4e4; }
</style>
"@

$starttime = Get-Date

[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
try {
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($server,$False,8530)
}
catch {
    Write-Error "Could not connect to WSUS: $_"
    return 1
}

$body = "";
$computerscope = new-object Microsoft.UpdateServices.Administration.ComputerTargetScope
$updatescope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

# Get failed updates
$updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Failed
$outlist = $wsus.GetUpdates($updatescope) | %{ $_.GetUpdateInstallationInfoPerComputerTarget($computerscope) | ?{ $_.UpdateInstallationState -eq "Failed"} |
 Select-Object @{L="Computer"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}},
 @{L="Update"; E={($wsus.GetUpdate([guid]$_.UpdateID)).Title}},
 @{L="Last error"; E={$compId = $_.ComputerTargetId; ($wsus.GetUpdateEventHistory((Get-Date).AddDays(-900), (Get-Date), [guid]$_.UpdateID, [Microsoft.UpdateServices.Administration.WsusEventSource]::Client, $null) | ?{ $_.ComputerId -eq $compId } | ?{ $_.IsError -eq $True -or $_.WsusEventId -eq "ClientDownloadCanceled" } | Select-Object -First 1).Message -replace "`n","" -replace "`r","" -replace 'Windows failed to install the following update with (error 0x[0-9a-f]*): .*$','$1'}},
 @{L="Last error time"; E={$compId = $_.ComputerTargetId; ($wsus.GetUpdateEventHistory((Get-Date).AddDays(-900), (Get-Date), [guid]$_.UpdateID, [Microsoft.UpdateServices.Administration.WsusEventSource]::Client, $null) | ?{ $_.ComputerId -eq $compId } | Select-Object -First 1).CreationDate.ToString("s")}},
 @{L="First error time"; E={$compId = $_.ComputerTargetId; ($wsus.GetUpdateEventHistory((Get-Date).AddDays(-900), (Get-Date), [guid]$_.UpdateID, [Microsoft.UpdateServices.Administration.WsusEventSource]::Client, $null) | ?{ $_.ComputerId -eq $compId } | Select-Object -Last 1).CreationDate.ToString("s")}}
 } | Sort-Object -Property Computer

$outfile = "C:\Program Files\Scripts\Data\WSUSReport-Failed-$(Get-Date -Format yyyy-MM-dd).csv"
$outlist | Export-Csv -NoTypeInformation -Path $outfile

$bodyraw = (Get-Content $outfile)
$body += ConvertFrom-Csv $bodyraw | ConvertTo-Html -Head $bodycss -Body "<h1>WSUS report</h1><h2>Failed updates</h2>" -PostContent "<h6>Generated on $(Get-Date) in $(((Get-Date) - $starttime).TotalMilliseconds / 1000) seconds</h6>" | Out-String

# Get update status report
$updatescope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled

# Get computer status report
$outlist = $wsus.GetSummariesPerComputerTarget($updatescope,$computerscope) | Select-Object @{L="Computer name"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}},@{L="Unk"; E={$_.UnknownCount}},@{L="Need"; E={$_.NotInstalledCount}},@{L="D/L"; E={$_.DownloadedCount}},@{L="Inst"; E={$_.InstalledCount}},@{L="Pend"; E={$_.InstalledPendingRebootCount}},@{L="Fail"; E={$_.FailedCount}} | Sort-Object -Property "Need" -Descending

$outfile3 = "C:\Program Files\Scripts\Data\WSUSReport-Computers-$(Get-Date -Format yyyy-MM-dd).csv"
$outlist | Export-Csv -NoTypeInformation -Path $outfile3

$bodyraw = (Get-Content $outfile3)
$body += ConvertFrom-Csv $bodyraw | ConvertTo-Html -Head $bodycss -Body "<h2>Computer status</h2>" -PostContent "<h6>Generated on $(Get-Date) in $(((Get-Date) - $starttime).TotalMilliseconds / 1000) seconds</h6>" | Out-String

# $outlist = wsus.GetUpdates($updatescope) | ?{ $_.IsApproved -eq $False -and $_.IsDeclined -eq $False } | Select ArrivalDate,Title,State,IsBeta,UpdateClassificationTitle,MsrcSeverity | ft
$outlist = $wsus.GetSummariesPerUpdate($updatescope,$computerscope) | Select-Object @{L="Downloaded"; E={$_.LastUpdated.ToString("s")}},@{L="Update"; E={($wsus.GetUpdate([guid]$_.UpdateID)).Title}},@{L="Unk"; E={$_.UnknownCount}},@{L="N/A"; E={$_.NotApplicableCount}},@{L="Need"; E={$_.NotInstalledCount}},@{L="D/L"; E={$_.DownloadedCount}},@{L="Inst"; E={$_.InstalledCount}},@{L="Pend"; E={$_.InstalledPendingRebootCount}},@{L="Fail"; E={$_.FailedCount}} | Sort-Object -Property Downloaded -Descending

$outfile2 = "C:\Program Files\Scripts\Data\WSUSReport-Updates-$(Get-Date -Format yyyy-MM-dd).csv"
$outlist | Export-Csv -NoTypeInformation -Path $outfile2

$bodyraw = (Get-Content $outfile2)
$body += ConvertFrom-Csv $bodyraw | ConvertTo-Html -Head $bodycss -Body "<h2>Update status</h2>" -PostContent "<h6>Generated on $(Get-Date) in $(((Get-Date) - $starttime).TotalMilliseconds / 1000) seconds</h6>" | Out-String

# and send the report
Send-MailMessage -To $recipient -From $mailfrom -Subject "WSUS report for day $(Get-Date -Format FileDate) on $(HOSTNAME)" -SmtpServer $mailserver -Body $body -BodyAsHtml -Priority Low -Attachments @($outfile, $outfile2, $outfile3)
