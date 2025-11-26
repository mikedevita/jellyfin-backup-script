# Author: Josh2kk
# Github : https://github.com/josh2kk/
# Tool : jellyfin_backup_manager.ps1
# Description: Backup and restore script for Jellyfin on Windows

param(
    [switch]$BackupOnly
)

Add-Type -AssemblyName System.Windows.Forms

# === CONFIGURATION ===
# Detect Jellyfin installation directory based on privilege
$ProgramDataPath = "$env:ProgramData\Jellyfin"
$LocalAppDataPath1 = Join-Path $env:LOCALAPPDATA "Jellyfin"
$LocalAppDataPath2 = Join-Path $env:LOCALAPPDATA "jellyfin"

if (Test-Path $ProgramDataPath) {
    $JellyfinDataPath = $ProgramDataPath
} elseif (Test-Path $LocalAppDataPath1) {
    $JellyfinDataPath = $LocalAppDataPath1
} elseif (Test-Path $LocalAppDataPath2) {
    $JellyfinDataPath = $LocalAppDataPath2
} else {
    Write-Host "Jellyfin data folder not found in expected locations."
    Write-Host "Please check if Jellyfin is installed and run this script as the same user running Jellyfin."
    exit 1
}

$DefaultBackupFolder = "$PSScriptRoot\Backups"
$TaskName = "JellyfinAutoBackup"

# === HEADER DISPLAY ===
Write-Host "============================================"
Write-Host "# Author: Josh2kk"
Write-Host "# Github : https://github.com/josh2kk/"
Write-Host "# Tool : jellyfin_backup_manager.ps1"
Write-Host "# Description: Backup and restore script for Jellyfin on Windows"
Write-Host "============================================`n"

function Select-FolderDialog {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Backup Destination Folder"
    $folderBrowser.ShowNewFolderButton = $true
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    }
    return $null
}

function Get-7zipPath {
    $7zipDir = Join-Path $PSScriptRoot "7zip"
    $7zaPath = Join-Path $7zipDir "7za.exe"
    $7zPath = Join-Path $7zipDir "7z.exe"
    
    # Check if 7zip already exists
    if (Test-Path $7zaPath) {
        return $7zaPath
    }
    if (Test-Path $7zPath) {
        return $7zPath
    }
    
    # 7zip not found, need to download it
    Write-Host "7zip not found. Downloading 7zip portable..."
    
    try {
        # Create 7zip directory
        if (-not (Test-Path $7zipDir)) {
            New-Item -ItemType Directory -Path $7zipDir | Out-Null
        }
        
        # Download 7zip extra package (contains 7za.exe)
        $7zipUrl = "https://www.7-zip.org/a/7z2301-extra.7z"
        $7zipArchive = Join-Path $env:TEMP "7z2301-extra.7z"
        
        Write-Host "Downloading from $7zipUrl..."
        Invoke-WebRequest -Uri $7zipUrl -OutFile $7zipArchive -UseBasicParsing
        
        # Extract 7zip archive
        # We'll try multiple extraction methods since we don't have 7zip yet
        $tempExtractDir = Join-Path $env:TEMP "7zip-extract"
        if (Test-Path $tempExtractDir) {
            Remove-Item -Path $tempExtractDir -Recurse -Force
        }
        New-Item -ItemType Directory -Path $tempExtractDir | Out-Null
        
        $extracted = $false
        
        # Method 1: Try Windows built-in tar (Windows 10 1903+)
        try {
            $tarOutput = & tar -xf $7zipArchive -C $tempExtractDir 2>&1
            if ($LASTEXITCODE -eq 0) {
                $extracted = $true
                Write-Host "Extracted using Windows tar."
            }
        } catch {
            # tar not available or failed
        }
        
        # Method 2: Try PowerShell Expand-Archive (might work for some 7z files)
        if (-not $extracted) {
            try {
                Expand-Archive -Path $7zipArchive -DestinationPath $tempExtractDir -Force -ErrorAction Stop
                $extracted = $true
                Write-Host "Extracted using PowerShell Expand-Archive."
            } catch {
                # Expand-Archive doesn't support 7z format
            }
        }
        
        if (-not $extracted) {
            throw "Could not extract 7zip archive automatically. Windows tar or PowerShell Expand-Archive failed."
        }
        
        # Find and copy 7za.exe or 7z.exe to our 7zip directory
        $7zaSource = Get-ChildItem -Path $tempExtractDir -Recurse -Filter "7za.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $7zaSource) {
            $7zaSource = Get-ChildItem -Path $tempExtractDir -Recurse -Filter "7z.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        }
        
        if ($7zaSource) {
            Copy-Item -Path $7zaSource.FullName -Destination $7zipDir -Force
            Write-Host "7zip downloaded and extracted successfully."
        } else {
            throw "7za.exe or 7z.exe not found in downloaded archive"
        }
        
        # Cleanup
        Remove-Item -Path $7zipArchive -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $tempExtractDir -Recurse -Force -ErrorAction SilentlyContinue
        
        # Verify the file exists now
        if (Test-Path $7zaPath) {
            return $7zaPath
        }
        if (Test-Path $7zPath) {
            return $7zPath
        }
        
        throw "7zip executable not found after extraction"
        
    } catch {
        Write-Host "Failed to download 7zip: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please manually download 7zip portable from https://www.7-zip.org/" -ForegroundColor Yellow
        Write-Host "Extract 7za.exe to: $7zipDir" -ForegroundColor Yellow
        throw "7zip is required but could not be downloaded automatically"
    }
}

function Stop-Jellyfin {
    Write-Host "Attempting to stop Jellyfin..."
    
    # Check if Jellyfin is running as a service
    if (Get-Service -Name "jellyfin" -ErrorAction SilentlyContinue) {
        Stop-Service -Name "jellyfin" -Force
        Write-Host "Jellyfin service stopped."
    } else {
        # If Jellyfin is running as a process, stop it
        $proc = Get-Process jellyfin -ErrorAction SilentlyContinue
        if ($proc) {
            $proc | Stop-Process -Force
            Write-Host "Jellyfin user process killed."
        } else {
            Write-Host "Jellyfin is not currently running."
        }
    }
    Start-Sleep -Seconds 3
}

function Start-Jellyfin {
    # Try starting the service first
    if (Get-Service -Name "jellyfin" -ErrorAction SilentlyContinue) {
        Start-Service -Name "jellyfin"
        Write-Host "Jellyfin service started."
    } else {
        # Otherwise, try starting it as a process (for user installs)
        $jellyfinPath = Join-Path $env:LOCALAPPDATA "Jellyfin\jellyfin.exe"
        if (Test-Path $jellyfinPath) {
            Start-Process -FilePath $jellyfinPath
            Write-Host "Jellyfin process started."
        } else {
            Write-Host "Jellyfin executable not found. Please start it manually."
        }
    }
}

function Create-Backup {
    Write-Host "`nStarting backup..."
    Stop-Jellyfin

    $choice = Read-Host "Choose backup location: [1] Script directory [2] Choose manually"
    if ($choice -eq '1') {
        $backupDir = $DefaultBackupFolder
    } elseif ($choice -eq '2') {
        $backupDir = Select-FolderDialog
        if (-not $backupDir) {
            Write-Host "No folder selected. Aborting."
            return
        }
    } else {
        Write-Host "Invalid choice."
        return
    }

    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFile = Join-Path $backupDir "JellyfinBackup_$timestamp.zip"

    try {
        Write-Host "Creating backup: $backupFile"
        
        # Get 7zip path (downloads if needed)
        $7zipExe = Get-7zipPath
        
        # Use 7zip to create ZIP archive
        # -tzip: use ZIP format
        # -mx5: compression level 5 (balanced)
        # -mm=Deflate: use deflate method for ZIP
        # -y: assume yes on all queries
        $arguments = @(
            "a",
            "-tzip",
            "-mx5",
            "-mm=Deflate",
            "-y",
            "`"$backupFile`"",
            "`"$JellyfinDataPath\*`""
        )
        
        $process = Start-Process -FilePath $7zipExe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Backup created successfully!"
        } else {
            throw "7zip exited with code $($process.ExitCode)"
        }
    } catch {
        Write-Host "Backup failed: $($_.Exception.Message)"
    }

    Start-Jellyfin
}

function Restore-Backup {
    Write-Host "`nRestore from Backup"
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "ZIP files (*.zip)|*.zip"
    $dialog.Title = "Select a Jellyfin Backup File"
    if ($dialog.ShowDialog() -eq "OK") {
        $zipPath = $dialog.FileName
        Write-Host "Restoring backup from: $zipPath"

        Stop-Jellyfin

        try {
            Remove-Item -Path "$JellyfinDataPath\*" -Recurse -Force
            
            # Get 7zip path (downloads if needed)
            $7zipExe = Get-7zipPath
            
            # Use 7zip to extract ZIP archive
            # x: extract with full paths
            # -o: output directory (no space after -o)
            # -y: assume yes on all queries
            $arguments = @(
                "x",
                "`"$zipPath`"",
                "-o`"$JellyfinDataPath`"",
                "-y"
            )
            
            $process = Start-Process -FilePath $7zipExe -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            
            if ($process.ExitCode -eq 0) {
                Write-Host "Restore completed successfully."
            } else {
                throw "7zip exited with code $($process.ExitCode)"
            }
        } catch {
            Write-Host "Restore failed: $($_.Exception.Message)"
        }

        Start-Jellyfin
    } else {
        Write-Host "Restore canceled by user."
    }
}

function Schedule-Backup {
    Write-Host "`nSchedule Automatic Backup"
    $freq = Read-Host "Choose backup frequency: [d]aily, [w]eekly, [m]onthly, [y]early"
    $intervalMap = @{
        'd' = 'DAILY'
        'w' = 'WEEKLY'
        'm' = 'MONTHLY'
        'y' = 'ONCE'
    }

    if (-not $intervalMap.ContainsKey($freq)) {
        Write-Host "Invalid frequency."
        return
    }

    $scheduleType = $intervalMap[$freq]
    $scriptPath = $MyInvocation.MyCommand.Definition

    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

    Register-ScheduledTask `
        -Action (New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -BackupOnly") `
        -Trigger (New-ScheduledTaskTrigger -$scheduleType -At 3am) `
        -TaskName $TaskName -Description "Jellyfin auto-backup" `
        -User "$env:UserName" -RunLevel Highest

    Write-Host "Scheduled $scheduleType backups at 3:00 AM."
}

# === MAIN ENTRY ===
if ($BackupOnly) {
    Create-Backup
    exit
}

while ($true) {
    Write-Host "`nJellyfin Backup Manager"
    Write-Host "1. Backup Jellyfin"
    Write-Host "2. Restore from Backup"
    Write-Host "3. Schedule Automatic Backups"
    Write-Host "4. Exit"
    $input = Read-Host "Select an option"

    switch ($input) {
        '1' { Create-Backup }
        '2' { Restore-Backup }
        '3' { Schedule-Backup }
        '4' { Write-Host "Exiting..."; exit }
        default { Write-Host "Invalid option. Try again." }
    }
}
