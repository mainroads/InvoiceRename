$pathtomonitor = Get-Location
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
                $fp = Join-Path $pathtomonitor $change.Name
                $creationDate = (Get-Item $fp).CreationTime
                $fileDate = $creationDate.ToString("yyyyMMdd")
                $folderDate = $creationDate.ToString("yyyyMM")
                
                $newPath = Join-Path $pathtomonitor $folderDate
                if(-not(Test-Path -Path $newPath)) {
                    New-Item -ItemType Directory -Path $newPath | Out-Null
                }

                Get-Item $fp | Rename-Item -NewName { $fileDate + " " + $_.Name } -PassThru |
                Move-Item -Destination $newPath -Force
            }
        }
        else {
            Write-Host "." -NoNewline
        }
    }
}
finally {
    $FileSystemWatcher.Dispose()
    Write-Host "My Watcher is done."
}
