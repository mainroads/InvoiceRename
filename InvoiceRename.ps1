$PathToMonitor = Get-Location
$timeout = 1000

try {
    $FileSystemWatcher = New-Object System.IO.FileSystemWatcher $pathtomonitor
    $FileSystemWatcher.IncludeSubdirectories = $true

    Write-Host "Monitoring content of $PathToMonitor"
    while ($true) {
        $change = $FileSystemWatcher.WaitForChanged('All', $timeout)
        if ($change.TimedOut -eq $false) {
            if ($change.ChangeType -eq 'Created') {
                Start-Sleep -Seconds 1
                $fp = Join-Path $PathToMonitor $change.Name
                $creationDate = (Get-Item $fp).CreationTime.ToString("yyyyMMdd")
                Get-Item $fp | Rename-Item -NewName { $creationDate + " " + $_.Name }
            }
        }
        else {
            Write-Host "." -NoNewline
        }
    }
}
finally {
    $FileSystemWatcher.Dispose()
}
