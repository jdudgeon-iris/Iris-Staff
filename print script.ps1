# Define time range
$endDate = Get-Date
$startDate = $endDate.AddDays(-31)

# Define output CSV path
$outputCsv = "C:\PrintLogsSummary.csv"

# Ensure the PrintService Operational log is enabled
wevtutil sl "Microsoft-Windows-PrintService/Operational" /e:true

# Get Event ID 307 logs in the date range
$logs = Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" |
    Where-Object {
        $_.TimeCreated -ge $startDate -and $_.TimeCreated -le $endDate -and $_.Id -eq 307
    }

# Parse logs
$parsedLogs = $logs | ForEach-Object {
    try {
        $xml = [xml]$_.ToXml()
        $event = $xml.Event
        $document = $event.UserData.DocumentPrinted

        [PSCustomObject]@{
            User      = $document.Param3
            Pages     = $document.Param8
            Printer   = $document.Param5
            Computer = $document.Param4
            Timestamp = $_.TimeCreated
        }
    } catch {
        Write-Warning "Failed to parse log entry at $($_.TimeCreated): $_"
    }
}

# Export to CSV
$parsedLogs | Export-Csv -Path $outputCsv -NoTypeInformation
Write-Host "Print log summary saved to $outputCsv"
