#
# PowerShell Script to Monitor a Folder and Organize Files (Enhanced Error-Proof Version)
#
# This script watches a specified directory and its subdirectories for newly created files.
# It uses a robust event-driven model to handle file creations reliably.
#
# This version DOES NOT require Outlook to be installed and is compatible with older PowerShell versions.
# - For .pdf files, it renames them with their creation date.
# - For .eml files, it parses the text content to find the email's sent date.
# - For .msg files, it uses the third-party MsgReader.dll library to read the email's sent date.
#
# PREREQUISITE for .msg files:
# 1. Download the MsgReader package from https://www.nuget.org/packages/MsgReader/
# 2. Rename the downloaded .nupkg file to .zip and extract 'MsgReader.dll' from the 'lib/net45' folder.
# 3. Place 'MsgReader.dll' in the same directory as this script.
# 4. Right-click 'MsgReader.dll', select Properties, and click the "Unblock" button or checkbox.
#

# Enable verbose debug output
$DebugPreference = "Continue"
$VerbosePreference = "Continue"

Write-Host "=== ENHANCED FILE MONITOR WITH DEBUG INFO ===" -ForegroundColor Cyan
Write-Host "Current PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Green
Write-Host "Current User: $($env:USERNAME)" -ForegroundColor Green
Write-Host "Script Start Time: $(Get-Date)" -ForegroundColor Green

# --- CONFIGURATION ---
# Load settings from Settings.json file
$settingsPath = Join-Path -Path $PSScriptRoot -ChildPath "Settings.json"
Write-Host "`n=== SETTINGS CONFIGURATION ===" -ForegroundColor Yellow
Write-Host "Settings file path: '$settingsPath'" -ForegroundColor Green

# Default settings
$defaultSettings = @{
    WatchFolderPath = $PSScriptRoot
}

# Load or create settings file
if (Test-Path $settingsPath) {
    try {
        Write-Host "Loading settings from existing file..." -ForegroundColor Green
        $settingsContent = Get-Content -Path $settingsPath -Raw -Encoding UTF8
        $settings = $settingsContent | ConvertFrom-Json
        Write-Host "Settings loaded successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR reading settings file: $_" -ForegroundColor Red
        Write-Host "Using default settings..." -ForegroundColor Yellow
        $settings = $defaultSettings
    }
} else {
    Write-Host "Settings file not found, creating with default values..." -ForegroundColor Yellow
    $settings = $defaultSettings
    try {
        $settings | ConvertTo-Json -Depth 3 | Set-Content -Path $settingsPath -Encoding UTF8
        Write-Host "Settings file created: '$settingsPath'" -ForegroundColor Green
    }
    catch {
        Write-Host "WARNING: Could not create settings file: $_" -ForegroundColor Yellow
    }
}

# Validate and set the watch folder path
if ($settings.WatchFolderPath -and (Test-Path $settings.WatchFolderPath)) {
    $pathtomonitor = Get-Item $settings.WatchFolderPath
    Write-Host "Using watch folder from settings: '$($pathtomonitor.Path)'" -ForegroundColor Green
} else {
    Write-Host "WARNING: Watch folder path in settings is invalid or doesn't exist" -ForegroundColor Yellow
    Write-Host "Falling back to script directory..." -ForegroundColor Yellow
    $pathtomonitor = Get-Item $PSScriptRoot
    # Update settings with the fallback path
    $settings.WatchFolderPath = $PSScriptRoot
    try {
        $settings | ConvertTo-Json -Depth 3 | Set-Content -Path $settingsPath -Encoding UTF8
        Write-Host "Settings file updated with fallback path" -ForegroundColor Green
    }
    catch {
        Write-Host "WARNING: Could not update settings file: $_" -ForegroundColor Yellow
    }
}

Write-Host "`n=== PATH CONFIGURATION ===" -ForegroundColor Yellow
Write-Host "Watch folder path: '$($pathtomonitor.Path)'" -ForegroundColor Green
Write-Host "PSScriptRoot: '$PSScriptRoot'" -ForegroundColor Green

# The script must be saved and run as a .ps1 file for the path detection to work.
if (-not $PSScriptRoot) {
    Write-Host "FATAL: This script must be saved as a .ps1 file and run from there." -ForegroundColor Red
    Write-Host "The automatic variable `$PSScriptRoot is not available in an interactive console." -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    return
}

# The script expects MsgReader.dll to be in the same folder as the script.
$scriptPath = $PSScriptRoot
$msgReaderDllPath = Join-Path -Path $scriptPath -ChildPath "MsgReader.dll"

Write-Host "Script Path: '$scriptPath'" -ForegroundColor Green
Write-Host "MsgReader DLL Path: '$msgReaderDllPath'" -ForegroundColor Green

# Use a global scope for the variable so it can be accessed inside the action block.
$global:msgReaderLoaded = $false
# Store the monitor path in global scope for reliable access in action block
$global:monitorPath = $pathtomonitor.Path

# Attempt to load the MsgReader DLL if it exists
Write-Host "`n=== MSG READER SETUP ===" -ForegroundColor Yellow
if (Test-Path $msgReaderDllPath) {
    Write-Host "MsgReader.dll found at: '$msgReaderDllPath'" -ForegroundColor Green
    try {
        # Unblock the file programmatically as a fallback.
        Unblock-File -Path $msgReaderDllPath -ErrorAction SilentlyContinue
        Add-Type -Path $msgReaderDllPath
        $global:msgReaderLoaded = $true
        Write-Host "MsgReader.dll loaded successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Error loading MsgReader.dll: $_" -ForegroundColor Red
        Write-Host "Please ensure the file is 'Unblocked' in its properties." -ForegroundColor Yellow
    }
} else {
    Write-Host "MsgReader.dll not found at '$msgReaderDllPath'. .msg file processing will be skipped." -ForegroundColor Yellow
}

# Test directory permissions
Write-Host "`n=== DIRECTORY PERMISSIONS TEST ===" -ForegroundColor Yellow
try {
    $testFile = Join-Path -Path $pathtomonitor.Path -ChildPath "test_permissions.tmp"
    "test" | Out-File -FilePath $testFile -Force
    Remove-Item -Path $testFile -Force
    Write-Host "Directory write permissions: OK" -ForegroundColor Green
} catch {
    Write-Host "Directory write permissions: FAILED - $_" -ForegroundColor Red
}

Write-Host "`nPreparing to monitor '$($pathtomonitor.Path)'..."

# --- HELPER FUNCTIONS ---
# Create functions in global scope so they can be accessed from the action script block
function global:Wait-ForFileAccess {
    param(
        [string]$FilePath,
        [int]$MaxWaitSeconds = 30,
        [int]$RetryIntervalSeconds = 1
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    Write-Host "[$timestamp] Waiting for file access: '$FilePath'" -ForegroundColor Yellow
    
    $waited = 0
    while ($waited -lt $MaxWaitSeconds) {
        try {
            # Try to open the file for reading and writing to ensure it's not locked
            $fileStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $fileStream.Close()
            Write-Host "[$timestamp] File is accessible after $waited seconds" -ForegroundColor Green
            return $true
        }
        catch {
            Start-Sleep -Seconds $RetryIntervalSeconds
            $waited += $RetryIntervalSeconds
        }
    }
    
    Write-Host "[$timestamp] WARNING: File may still be locked after $MaxWaitSeconds seconds" -ForegroundColor Yellow
    return $false
}

function global:Get-SafeFileName {
    param([string]$FileName)
    
    # Remove or replace characters that are invalid in file names
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $safeName = $FileName
    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace($char, '_')
    }
    return $safeName
}

function global:Move-FileWithRetry {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 2
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            Write-Host "[$timestamp] Move attempt $i of $MaxRetries" -ForegroundColor Yellow
            
            # Ensure destination directory exists
            $destinationDir = Split-Path -Path $DestinationPath -Parent
            if (-not (Test-Path -Path $destinationDir)) {
                Write-Host "[$timestamp] Creating destination directory: '$destinationDir'" -ForegroundColor Yellow
                New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null
            }
            
            # Check if destination file exists and handle accordingly
            if (Test-Path -Path $DestinationPath) {
                $counter = 1
                $fileInfo = [System.IO.FileInfo]$DestinationPath
                $baseName = $fileInfo.BaseName
                $extension = $fileInfo.Extension
                $directory = $fileInfo.Directory.FullName
                
                do {
                    $DestinationPath = Join-Path -Path $directory -ChildPath "$baseName`_$counter$extension"
                    $counter++
                } while (Test-Path -Path $DestinationPath)
                
                Write-Host "[$timestamp] Destination renamed to avoid conflict: '$DestinationPath'" -ForegroundColor Yellow
            }
            
            # Perform the move
            Move-Item -Path $SourcePath -Destination $DestinationPath -Force
            
            # Verify the move was successful
            if (Test-Path -Path $DestinationPath -and -not (Test-Path -Path $SourcePath)) {
                Write-Host "[$timestamp] SUCCESS: File moved to '$DestinationPath'" -ForegroundColor Green
                return $DestinationPath
            } else {
                throw "Move operation did not complete successfully"
            }
        }
        catch {
            Write-Host "[$timestamp] Move attempt $i failed: $_" -ForegroundColor Red
            if ($i -lt $MaxRetries) {
                Write-Host "[$timestamp] Retrying in $RetryDelaySeconds seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $RetryDelaySeconds
            } else {
                Write-Host "[$timestamp] All move attempts failed" -ForegroundColor Red
                throw $_
            }
        }
    }
}

# --- ACTION SCRIPT BLOCK ---
# This block of code runs every time a "Created" event is detected by the FileSystemWatcher.
$action = {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    Write-Host "`n[$timestamp] === FILE EVENT TRIGGERED ===" -ForegroundColor Cyan
    
    try {
        # The $event object contains information about the event that was triggered.
        # We get the full, unambiguous path of the new file from the event arguments.
        $fullPath = $event.SourceEventArgs.FullPath
        $name = $event.SourceEventArgs.Name
        $changeType = $event.SourceEventArgs.ChangeType
        
        Write-Host "[$timestamp] Event Type: $changeType" -ForegroundColor Green
        Write-Host "[$timestamp] File Name: '$name'" -ForegroundColor Green
        Write-Host "[$timestamp] Full Path: '$fullPath'" -ForegroundColor Green
        
        # Check if path exists immediately
        $pathExists = Test-Path -Path $fullPath
        Write-Host "[$timestamp] Path exists immediately: $pathExists" -ForegroundColor $(if($pathExists){"Green"}else{"Red"})
        
        if ($pathExists) {
            $itemType = if (Test-Path -Path $fullPath -PathType Leaf) { "File" } else { "Directory" }
            Write-Host "[$timestamp] Item type: $itemType" -ForegroundColor Green
            
            if ($itemType -eq "File") {
                $fileInfo = Get-Item $fullPath -ErrorAction SilentlyContinue
                if ($fileInfo) {
                    Write-Host "[$timestamp] File size: $($fileInfo.Length) bytes" -ForegroundColor Green
                    Write-Host "[$timestamp] File creation time: $($fileInfo.CreationTime)" -ForegroundColor Green
                    Write-Host "[$timestamp] File last write time: $($fileInfo.LastWriteTime)" -ForegroundColor Green
                }
            }
        }

        # Wait for file to be fully written and accessible
        Write-Host "[$timestamp] Waiting for file to stabilize..." -ForegroundColor Yellow
        Wait-ForFileAccess -FilePath $fullPath -MaxWaitSeconds 10

        # Verify the created item is a file and still exists before proceeding.
        if (-not(Test-Path -Path $fullPath -PathType Leaf)) {
            Write-Host "[$timestamp] SKIPPING: '$name' - not a file or was removed" -ForegroundColor Red
            if (Test-Path -Path $fullPath) {
                Write-Host "[$timestamp] Item exists but is not a file (probably a directory)" -ForegroundColor Yellow
            } else {
                Write-Host "[$timestamp] Item was removed or never existed" -ForegroundColor Red
            }
            return # Exit the action block for this event.
        }

        Write-Host "[$timestamp] File verification passed - proceeding with processing" -ForegroundColor Green        # Get the file item and its extension.
        $fileItem = Get-Item $fullPath -ErrorAction Stop
        $extension = $fileItem.Extension.ToLower()
        $emailDate = $null
        
        Write-Host "[$timestamp] File extension: '$extension'" -ForegroundColor Green
        Write-Host "[$timestamp] File base name: '$($fileItem.BaseName)'" -ForegroundColor Green

        # Skip .ini files
        if ($extension -eq '.ini') {
            Write-Host "[$timestamp] IGNORE: Skipping .ini file: $($fileItem.Name)" -ForegroundColor Yellow
            return # Exit the action block for .ini files.
        }

        # Use if/elseif/else to handle different file types.
        if ($extension -eq '.pdf') {
            Write-Host "[$timestamp] Processing PDF file: $($fileItem.Name)" -ForegroundColor Magenta
            $emailDate = $fileItem.CreationTime
            Write-Host "[$timestamp] Using PDF creation time: $emailDate" -ForegroundColor Green
        }
        elseif ($extension -eq '.eml') {
            Write-Host "[$timestamp] Processing EML file: $($fileItem.Name)" -ForegroundColor Magenta
            try {
                Write-Host "[$timestamp] Reading EML file content..." -ForegroundColor Yellow
                # Read the file content and find the 'Date:' header using regex.
                $content = Get-Content -Path $fullPath -Raw -Encoding UTF8 -ErrorAction Stop
                Write-Host "[$timestamp] File content length: $($content.Length) characters" -ForegroundColor Green
                
                # Show first few lines for debugging
                $lines = $content -split "`r?`n" | Where-Object { $_.Trim() -ne "" } | Select-Object -First 10
                Write-Host "[$timestamp] First 10 non-empty lines of EML file:" -ForegroundColor Yellow
                for ($i = 0; $i -lt $lines.Count; $i++) {
                    $linePreview = $lines[$i].Substring(0, [Math]::Min(100, $lines[$i].Length))
                    Write-Host "[$timestamp]   Line $($i+1): $linePreview" -ForegroundColor Gray
                }
                
                $match = [regex]::Match($content, '(?im)^Date:\s*(.+)$')
                if ($match.Success) {
                    $dateString = $match.Groups[1].Value.Trim()
                    Write-Host "[$timestamp] Found Date header: '$dateString'" -ForegroundColor Green
                    # Parse the extracted date string into a DateTime object.
                    $emailDate = [datetime]::Parse($dateString)
                    Write-Host "[$timestamp] Parsed email date: $emailDate" -ForegroundColor Green
                } else {
                    Write-Host "[$timestamp] Could not find 'Date:' header in EML file. Searching for alternative date headers..." -ForegroundColor Yellow
                    
                    # Try alternative date headers
                    $alternativePatterns = @(
                        '(?im)^Sent:\s*(.+)$',
                        '(?im)^Delivery-Date:\s*(.+)$',
                        '(?im)^Received:\s*.*;\s*(.+)$'
                    )
                    
                    foreach ($pattern in $alternativePatterns) {
                        $altMatch = [regex]::Match($content, $pattern)
                        if ($altMatch.Success) {
                            $dateString = $altMatch.Groups[1].Value.Trim()
                            Write-Host "[$timestamp] Found alternative date: '$dateString'" -ForegroundColor Green
                            try {
                                $emailDate = [datetime]::Parse($dateString)
                                Write-Host "[$timestamp] Successfully parsed alternative date: $emailDate" -ForegroundColor Green
                                break
                            } catch {
                                Write-Host "[$timestamp] Could not parse alternative date: $_" -ForegroundColor Yellow
                            }
                        }
                    }
                    
                    if (-not $emailDate) {
                        Write-Host "[$timestamp] No valid date found in headers. Using file creation time." -ForegroundColor Yellow
                        $emailDate = $fileItem.CreationTime
                    }
                }
            } catch {
                 Write-Host "[$timestamp] ERROR parsing EML file '$($fileItem.Name)': $_. Using file creation time." -ForegroundColor Red
                 $emailDate = $fileItem.CreationTime
            }
        }
        elseif ($extension -eq '.msg') {
            Write-Host "[$timestamp] Processing MSG file: $($fileItem.Name)" -ForegroundColor Magenta
            if (-not $Global:msgReaderLoaded) {
                 Write-Host "[$timestamp] SKIP: MSG file because MsgReader.dll is not loaded." -ForegroundColor Yellow
                 return
            }
            
            $fileStream = $null
            $reader = $null
            try {
                Write-Host "[$timestamp] Creating file stream for MSG file..." -ForegroundColor Yellow
                # Create a file stream to pass to the MsgReader, which requires a stream object.
                $fileStream = New-Object System.IO.FileStream($fullPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
                Write-Host "[$timestamp] Creating MsgReader instance..." -ForegroundColor Yellow
                # Use the MsgReader library to open the stream.
                $reader = New-Object MsgReader.Reader($fileStream)
                
                # Check if the Message object or SentOn property is null.
                if ($null -ne $reader.Message -and $null -ne $reader.Message.SentOn) {
                    # Get the date the email was sent by accessing the Message property.
                    $emailDate = $reader.Message.SentOn
                    Write-Host "[$timestamp] MSG file sent date: $emailDate" -ForegroundColor Green
                } else {
                    Write-Host "[$timestamp] Could not read 'SentOn' date from MSG file. Using file creation time instead." -ForegroundColor Yellow
                    $emailDate = $fileItem.CreationTime
                }
            } catch {
                Write-Host "[$timestamp] ERROR processing MSG file '$($fileItem.Name)': $_" -ForegroundColor Red
                Write-Host "[$timestamp] Falling back to using file creation time." -ForegroundColor Yellow
                $emailDate = $fileItem.CreationTime
            } finally {
                # Ensure all disposable objects are closed to release file locks.
                if ($reader) { 
                    Write-Host "[$timestamp] Disposing MsgReader..." -ForegroundColor Yellow
                    $reader.Dispose() 
                }
                if ($fileStream) { 
                    Write-Host "[$timestamp] Disposing file stream..." -ForegroundColor Yellow
                    $fileStream.Dispose() 
                }
            }
        }
        else {
            Write-Host "[$timestamp] IGNORE: Unhandled file extension '$extension' for file: $($fileItem.Name)" -ForegroundColor Yellow
            return # Exit the action block for unhandled files.
        }

        # --- RENAME AND MOVE LOGIC ---
        Write-Host "[$timestamp] === RENAME AND MOVE LOGIC ===" -ForegroundColor Cyan
        Write-Host "[$timestamp] Email/File date determined: '$emailDate'" -ForegroundColor Green
        
        if ($emailDate) {
            $fileDate = $emailDate.ToString("yyyyMMdd")
            $folderDate = $emailDate.ToString("yyyyMM")
            Write-Host "[$timestamp] File date string: '$fileDate'" -ForegroundColor Green
            Write-Host "[$timestamp] Folder date string: '$folderDate'" -ForegroundColor Green
            
            # Use the global monitor path for reliability
            $monitorPath = $Global:monitorPath
            Write-Host "[$timestamp] Monitor path from global: '$monitorPath'" -ForegroundColor Green
            
            # Validate monitor path exists
            if (-not (Test-Path -Path $monitorPath)) {
                Write-Host "[$timestamp] ERROR: Monitor path does not exist: '$monitorPath'" -ForegroundColor Red
                return
            }
            
            # The destination folder is always created in the root of the monitored path.
            $newPath = Join-Path -Path $monitorPath -ChildPath $folderDate
            Write-Host "[$timestamp] Destination folder: '$newPath'" -ForegroundColor Green
            
            # If the destination folder doesn't exist, create it.
            if (-not(Test-Path -Path $newPath)) {
                try {
                    Write-Host "[$timestamp] Creating directory: $newPath" -ForegroundColor Yellow
                    New-Item -ItemType Directory -Path $newPath -Force | Out-Null
                    Write-Host "[$timestamp] Directory created successfully" -ForegroundColor Green
                } catch {
                    Write-Host "[$timestamp] ERROR creating directory: $_" -ForegroundColor Red
                    return
                }
            } else {
                Write-Host "[$timestamp] Destination directory already exists" -ForegroundColor Green
            }            # Define the new name for the file (e.g., '20230422 MyDocument.pdf').
            $safeName = Get-SafeFileName -FileName $fileItem.Name
            
            # Check if the filename already starts with a date in yyyyMMdd format
            $datePattern = '^\d{8}\s'
            if ($safeName -match $datePattern) {
                Write-Host "[$timestamp] File already has date prefix, using original name: '$safeName'" -ForegroundColor Green
                $newName = $safeName
            } else {
                Write-Host "[$timestamp] Adding date prefix to filename" -ForegroundColor Green
                $newName = "$fileDate $safeName"
            }
            
            $destinationFile = Join-Path $newPath $newName
            
            Write-Host "[$timestamp] New file name: '$newName'" -ForegroundColor Green
            Write-Host "[$timestamp] Full destination path: '$destinationFile'" -ForegroundColor Green
            
            # Move the file using the retry function
            try {
                $finalPath = Move-FileWithRetry -SourcePath $fullPath -DestinationPath $destinationFile -MaxRetries 3
                Write-Host "[$timestamp] SUCCESS: File successfully moved to '$finalPath'" -ForegroundColor Green
            } catch {
                Write-Host "[$timestamp] FINAL ERROR: Could not move file after all retries: $_" -ForegroundColor Red
            }
            
        } else {
            Write-Host "[$timestamp] SKIP: No valid date determined, file will not be moved" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[$timestamp] CRITICAL ERROR in action block: $_" -ForegroundColor Red
        Write-Host "[$timestamp] Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    }
    finally {
        Write-Host "[$timestamp] === EVENT PROCESSING COMPLETE ===" -ForegroundColor Cyan
    }
}

# --- SETUP AND RUN ---
Write-Host "`n=== FILESYSTEM WATCHER SETUP ===" -ForegroundColor Yellow

# Test if the path is valid
if (-not (Test-Path -Path $pathtomonitor.Path)) {
    Write-Host "ERROR: Monitor path does not exist: '$($pathtomonitor.Path)'" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    return
}

# Set up the FileSystemWatcher object.
try {
    $FileSystemWatcher = New-Object System.IO.FileSystemWatcher $pathtomonitor.Path
    $FileSystemWatcher.IncludeSubdirectories = $true
    $FileSystemWatcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::CreationTime -bor [System.IO.NotifyFilters]::Size
    
    Write-Host "FileSystemWatcher created successfully" -ForegroundColor Green
    Write-Host "Path: '$($FileSystemWatcher.Path)'" -ForegroundColor Green
    Write-Host "Include Subdirectories: $($FileSystemWatcher.IncludeSubdirectories)" -ForegroundColor Green
    Write-Host "Notify Filter: $($FileSystemWatcher.NotifyFilter)" -ForegroundColor Green
    
    # This is crucial for the event-driven model to work.
    $FileSystemWatcher.EnableRaisingEvents = $true
    Write-Host "EnableRaisingEvents: $($FileSystemWatcher.EnableRaisingEvents)" -ForegroundColor Green

    # Register the event subscription. PowerShell will now listen for "Created" events
    # and run the $action script block automatically when one occurs.
    $subscriber = Register-ObjectEvent -InputObject $FileSystemWatcher -EventName "Created" -Action $action
    Write-Host "Event subscription registered: $($subscriber.Name)" -ForegroundColor Green

} catch {
    Write-Host "ERROR: Failed to set up FileSystemWatcher: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    return
}

# Show additional debug info
Write-Host "`n=== MONITORING STATUS ===" -ForegroundColor Yellow
Write-Host "Current directory contents:" -ForegroundColor Green
Get-ChildItem -Path $pathtomonitor.Path | ForEach-Object {
    Write-Host "  $($_.Name) ($($_.GetType().Name))" -ForegroundColor Gray
}

Write-Host "`n=== READY FOR MONITORING ===" -ForegroundColor Cyan
Write-Host "Monitoring path: '$($pathtomonitor.Path)'" -ForegroundColor Green
Write-Host "Supported file types: .pdf, .eml, .msg" -ForegroundColor Green
Write-Host "Current time: $(Get-Date)" -ForegroundColor Green
Write-Host "`nTo test: Drag and drop a .eml file into the monitored folder" -ForegroundColor Yellow
Write-Host "Press CTRL+C to stop monitoring" -ForegroundColor Cyan

try {
    # This loop keeps the script running so it can listen for events.
    # The actual work happens in the $action block when an event is fired.
    $eventCount = 0
    while ($true) {
        $result = Wait-Event -Timeout 10
        if ($result) {
            $eventCount++
            Write-Host "`nProcessed event #$eventCount at $(Get-Date)" -ForegroundColor Cyan
        } else {
            # Timeout occurred, show we're still alive
            Write-Host "." -NoNewline -ForegroundColor Gray
        }
    }
}
catch [System.Management.Automation.PipelineStoppedException] {
    Write-Host "`nMonitoring stopped by user (CTRL+C)" -ForegroundColor Yellow
}
catch {
    # Catch any script-terminating errors and display them.
    Write-Host "`n"
    Write-Host "An unexpected error occurred and the script has to stop." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    # Pause the script so the user can read the error message before the window closes.
    Read-Host "Press Enter to exit..."
}
finally {
    # This block always runs, ensuring resources are cleaned up when the script is stopped (e.g., by pressing CTRL+C).
    Write-Host "`n=== CLEANUP ===" -ForegroundColor Yellow
    Write-Host "Stopping monitoring..."
      if ($subscriber) {
        try {
            # Fix: Use -SourceIdentifier instead of -SubscriberIdentifier
            Unregister-Event -SourceIdentifier $subscriber.Name -ErrorAction Stop
            Write-Host "Event subscription unregistered" -ForegroundColor Green
        } catch {
            Write-Host "Error unregistering event: $_" -ForegroundColor Red
            # Try alternative cleanup method
            try {
                Get-EventSubscriber | Where-Object { $_.SourceObject -eq $FileSystemWatcher } | Unregister-Event
                Write-Host "Event subscription cleaned up using alternative method" -ForegroundColor Green
            } catch {
                Write-Host "Alternative cleanup also failed: $_" -ForegroundColor Red
            }
        }
    }
    
    if ($FileSystemWatcher) {
        try {
            $FileSystemWatcher.EnableRaisingEvents = $false
            $FileSystemWatcher.Dispose()
            Write-Host "FileSystemWatcher disposed" -ForegroundColor Green
        } catch {
            Write-Host "Error disposing FileSystemWatcher: $_" -ForegroundColor Red
        }
    }
    
    Write-Host "Cleanup complete. Script ended at $(Get-Date)" -ForegroundColor Green
}