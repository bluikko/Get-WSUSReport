$updateServer = "wsus.example.com"
$updateServerSSL = $False
$updateServerPort = 8530

$reportRecipient = "WSUS Master <wsus.master@example.com>"
$reportFrom = "WSUS report $(HOSTNAME) <wsusreport@example.com>"
$mailServer = "mail.example.com"

$reportFolder = "C:\Program Files\Scripts\Data"

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

$startTime = Get-Date

[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
try {
    $wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($updateServer,$updateServerSSL,$updateServerPort)
}
catch {
    Write-Error "Could not connect to WSUS: $_"
    return 1
}

$body = "";
$computerScope = new-object Microsoft.UpdateServices.Administration.ComputerTargetScope
$updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

# Get failed updates report
$updateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::Failed
$outList = $wsus.GetUpdates($updateScope) | %{ $_.GetUpdateInstallationInfoPerComputerTarget($computerScope) | ?{ $_.UpdateInstallationState -eq "Failed"} |
 Select-Object @{L="Computer"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}},
 @{L="Update"; E={($wsus.GetUpdate([guid]$_.UpdateID)).Title}},
 @{L="Last error"; E={$compId = $_.ComputerTargetId; ($wsus.GetUpdateEventHistory((Get-Date).AddDays(-900), (Get-Date), [guid]$_.UpdateID, [Microsoft.UpdateServices.Administration.WsusEventSource]::Client, $null) | ?{ $_.ComputerId -eq $compId } | ?{ $_.IsError -eq $True -or $_.WsusEventId -eq "ClientDownloadCanceled" } | Select-Object -First 1).Message -replace "`n","" -replace "`r","" -replace 'Windows failed to install the following update with (error 0x[0-9a-f]*): .*$','$1'}},
 @{L="Last error time"; E={$compId = $_.ComputerTargetId; ($wsus.GetUpdateEventHistory((Get-Date).AddDays(-900), (Get-Date), [guid]$_.UpdateID, [Microsoft.UpdateServices.Administration.WsusEventSource]::Client, $null) | ?{ $_.ComputerId -eq $compId } | Select-Object -First 1).CreationDate.ToString("s")}},
 @{L="First error time"; E={$compId = $_.ComputerTargetId; ($wsus.GetUpdateEventHistory((Get-Date).AddDays(-900), (Get-Date), [guid]$_.UpdateID, [Microsoft.UpdateServices.Administration.WsusEventSource]::Client, $null) | ?{ $_.ComputerId -eq $compId } | Select-Object -Last 1).CreationDate.ToString("s")}}
 } | Sort-Object -Property Computer

$outFile = "$($reportFolder)\WSUSReport-Failed-$(Get-Date -Format yyyy-MM-dd).csv"
$outList | Export-Csv -NoTypeInformation -Path $outFile

$bodyRaw = (Get-Content $outFile)
$body += ConvertFrom-Csv $bodyRaw | ConvertTo-Html -Head $bodycss -Body "<h1>WSUS report</h1><h2>Failed updates</h2>" -PostContent "<h6>Generated on $(Get-Date) in $(((Get-Date) - $startTime).TotalMilliseconds / 1000) seconds</h6>" | Out-String

# Get computer status report
$startTime = Get-Date
$updateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled

$outList = $wsus.GetSummariesPerComputerTarget($updateScope,$computerScope) | Select-Object @{L="Computer name"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).FullDomainName}},@{L="OS Name"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).OSDescription -replace "Windows ","" -replace " Edition","" -replace " installation",""}},@{L="OS Version"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).ClientVersion}},@{L="Sync Time"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).LastSyncTime.ToString("s")}},@{L="Report Time"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).LastReportedStatusTime.ToString("s")}},@{L="Sync Result"; E={($wsus.GetComputerTarget([guid]$_.ComputerTargetId)).LastSyncResult}},@{L="Unk"; E={$_.UnknownCount}},@{L="Need"; E={$_.NotInstalledCount}},@{L="D/L"; E={$_.DownloadedCount}},@{L="Inst"; E={$_.InstalledCount}},@{L="Pend"; E={$_.InstalledPendingRebootCount}},@{L="Fail"; E={$_.FailedCount}} | Sort-Object -Property "Need" -Descending

$outFile3 = "$($reportFolder)\WSUSReport-Computers-$(Get-Date -Format yyyy-MM-dd).csv"
$outList | Export-Csv -NoTypeInformation -Path $outFile3

$bodyRaw = (Get-Content $outFile3)
$body += ConvertFrom-Csv $bodyRaw | ConvertTo-Html -Head $bodycss -Body "<h2>Computer status</h2>" -PostContent "<h6>Generated on $(Get-Date) in $(((Get-Date) - $startTime).TotalMilliseconds / 1000) seconds</h6>" | Out-String

# Get update status report
$startTime = Get-Date

# $outList = wsus.GetUpdates($updateScope) | ?{ $_.IsApproved -eq $False -and $_.IsDeclined -eq $False } | Select ArrivalDate,Title,State,IsBeta,UpdateClassificationTitle,MsrcSeverity | ft
$outList = $wsus.GetSummariesPerUpdate($updateScope,$computerScope) | Select-Object @{L="Downloaded"; E={$_.LastUpdated.ToString("s")}},@{L="Update"; E={($wsus.GetUpdate([guid]$_.UpdateID)).Title}},@{L="Unk"; E={$_.UnknownCount}},@{L="N/A"; E={$_.NotApplicableCount}},@{L="Need"; E={$_.NotInstalledCount}},@{L="D/L"; E={$_.DownloadedCount}},@{L="Inst"; E={$_.InstalledCount}},@{L="Pend"; E={$_.InstalledPendingRebootCount}},@{L="Fail"; E={$_.FailedCount}} | Sort-Object -Property Downloaded -Descending

$outFile2 = "$($reportFolder)\WSUSReport-Updates-$(Get-Date -Format yyyy-MM-dd).csv"
$outList | Export-Csv -NoTypeInformation -Path $outFile2

$bodyRaw = (Get-Content $outFile2)
$body += ConvertFrom-Csv $bodyRaw | ConvertTo-Html -Head $bodycss -Body "<h2>Update status</h2>" -PostContent "<h6>Generated on $(Get-Date) in $(((Get-Date) - $startTime).TotalMilliseconds / 1000) seconds</h6>" | Out-String

# and send the report
Send-MailMessage -To $reportRecipient -From $reportFrom -Subject "WSUS report for day $(Get-Date -Format FileDate) on $(HOSTNAME)" -SmtpServer $mailServer -Body $body -BodyAsHtml -Priority Low -Attachments @($outFile, $outFile2, $outFile3)
