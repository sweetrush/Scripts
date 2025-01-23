# Add some color functions
function Write-ColorOutput {
    param (
        [string]$Message,
        [string]$ForegroundColor = "White",
        [string]$BackgroundColor = "Black",
        [switch]$NoNewLine
    )
    $previousForeground = $host.UI.RawUI.ForegroundColor
    $previousBackground = $host.UI.RawUI.BackgroundColor
    
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    $host.UI.RawUI.BackgroundColor = $BackgroundColor
    
    if ($NoNewLine) {
        Write-Host $Message -NoNewLine
    } else {
        Write-Host $Message
    }
    
    $host.UI.RawUI.ForegroundColor = $previousForeground
    $host.UI.RawUI.BackgroundColor = $previousBackground
}

# Banner function
function Show-Banner {
    Write-ColorOutput "`n=============================================" "Cyan"
    Write-ColorOutput "          Desktop File Organizer" "Yellow"
    Write-ColorOutput "=============================================" "Cyan"
    Write-ColorOutput "Starting organization process...`n" "Green"
}

# Script configuration
$config = @{
    DesktopPath = [Environment]::GetFolderPath("Desktop")
    LogPath = Join-Path -Path $env:TEMP -ChildPath "DesktopOrganizer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    CreateBackup = $true
    ShowProgress = $true
    UnmovedFilesReport = Join-Path -Path $env:TEMP -ChildPath "UnmovedFiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
}

# Enhanced logging function with colors
function Write-ToLog {
    param(
        [string]$Message,
        [string]$Type = "INFO"  # INFO, SUCCESS, WARNING, ERROR
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] $Type : $Message"
    Add-Content -Path $config.LogPath -Value $logMessage
    
    switch ($Type) {
        "INFO"    { Write-ColorOutput $logMessage "Cyan" }
        "SUCCESS" { Write-ColorOutput $logMessage "Green" }
        "WARNING" { Write-ColorOutput $logMessage "Yellow" }
        "ERROR"   { Write-ColorOutput $logMessage "Red" }
        default   { Write-ColorOutput $logMessage "White" }
    }
}

# Function to create backup with colorful output
function Backup-Desktop {
    Write-ColorOutput "`n[Backup] " "Magenta" -NoNewLine
    Write-ColorOutput "Creating desktop backup..." "White"
    $backupPath = Join-Path -Path $env:TEMP -ChildPath "DesktopBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Copy-Item -Path $config.DesktopPath -Destination $backupPath -Recurse
    Write-ToLog "Backup created at: $backupPath" "SUCCESS"
    return $backupPath
}

# Function to check if a file is locked
function Test-FileLock {
    param (
        [parameter(Mandatory=$true)]
        [string]$Path
    )
    $locked = $false
    $file = $null
    try {
        $file = [IO.File]::Open($Path, 'Open', 'Read', 'None')
    }
    catch {
        $locked = $true
    }
    finally {
        if ($file) {
            $file.Close()
        }
    }
    return $locked
}

# Function to move items with retry logic
function Move-ItemWithRetry {
    param (
        [string]$Source,
        [string]$Destination,
        [int]$MaxAttempts = 3
    )
    
    $attempt = 1
    while ($attempt -le $MaxAttempts) {
        if (Test-FileLock -Path $Source) {
            Write-ToLog "File is locked: $Source" "WARNING"
            Start-Sleep -Seconds 1
            $attempt++
            continue
        }
        
        try {
            Move-Item -Path $Source -Destination $Destination -ErrorAction Stop
            return $true
        }
        catch {
            if ($attempt -eq $MaxAttempts) {
                Write-ToLog "Failed to move $Source after $MaxAttempts attempts: $_" "ERROR"
                return $false
            }
            Start-Sleep -Seconds 1
            $attempt++
        }
    }
}

# Function to get unmoved files
function Get-UnmovedFiles {
    param (
        [string]$Path
    )
    
    $excludedFolders = @("TAEP", "BWS", "WWS", "CW", "PRP", "Documents", "Images", 
                         "Audio", "Videos", "Archives", "Executables", "Shortcuts", 
                         "Others", "Folders")
    
    Get-ChildItem -Path $Path -File | 
        Where-Object { $_.Directory.Name -eq (Split-Path -Path $Path -Leaf) }
}

# Categories definition
$Categories = @{
    PrefixBased = @{
        TAEP = @{
            Pattern = "TAEP_*"
            Description = "TAEP Files"
            Color = "Yellow"
        }
        BWS = @{
            Pattern = "BWS_*"
            Description = "BWS Files"
            Color = "Green"
        }
        WWS = @{
            Pattern = "WWS_*"
            Description = "WWS Files"
            Color = "Cyan"
        }
        CW = @{
            Pattern = "CW_*"
            Description = "CW Files"
            Color = "Magenta"
        }
        PRP = @{
            Pattern = "PRP_*"
            Description = "PRP Files"
            Color = "Blue"
        }
    }
    ExtensionBased = @{
        Documents = @{
            Pattern = "*.txt,*.docx,*.pdf,*.csv,*.vsdx,*.pub,*.xlsx,*.doc,*.rtf,*.pptx,*.xls,*.odt,*.ods,*.odp,*.pages,*.numbers"
            Description = "Document Files"
            Color = "White"
        }
        Images = @{
            Pattern = "*.jpg,*.jpeg,*.png,*.gif,*.bmp,*.tiff,*.ico,*.svg,*.raw,*.cr2,*.nef,*.webp,*.heic"
            Description = "Image Files"
            Color = "Green"
        }
        Audio = @{
            Pattern = "*.wav,*.mp3,*.flac,*.m4a,*.aac,*.wma,*.ogg,*.mid,*.midi,*.aiff"
            Description = "Audio Files"
            Color = "Yellow"
        }
        Videos = @{
            Pattern = "*.mp4,*.avi,*.mkv,*.mov,*.wmv,*.flv,*.webm,*.m4v,*.mpg,*.mpeg,*.3gp"
            Description = "Video Files"
            Color = "Magenta"
        }
        Archives = @{
            Pattern = "*.zip,*.7z,*.rar,*.tar,*.gz,*.bz2,*.xz,*.iso,*.tgz"
            Description = "Archive Files"
            Color = "DarkYellow"
        }
        Executables = @{
            Pattern = "*.exe,*.msi,*.bat,*.cmd,*.ps1,*.vbs,*.jar,*.appx,*.msix"
            Description = "Executable Files"
            Color = "Red"
        }
        Shortcuts = @{
            Pattern = "*.lnk,*.url,*.desktop"
            Description = "Shortcut Files"
            Color = "Blue"
        }
        Others = @{
            Pattern = "*.*"
            Description = "Other Files"
            Color = "Gray"
        }
    }
}

# Main script execution
try {
    Show-Banner
    
   

    # Calculate total items for progress
    $totalItems = (Get-ChildItem -Path $config.DesktopPath -File).Count
    $processedItems = 0
    $successfulMoves = 0
    $failedMoves = 0

    # Create directories with colorful output
    Write-ColorOutput "`n[Creating Directories]" "Yellow"
    foreach ($categoryType in $Categories.Keys) {
        foreach ($category in $Categories[$categoryType].Keys) {
            $categoryPath = Join-Path -Path $config.DesktopPath -ChildPath $category
            if (!(Test-Path -Path $categoryPath -PathType Container)) {
                New-Item -Path $categoryPath -ItemType Directory -Force | Out-Null
                Write-ColorOutput "  ► Created: $category" $Categories[$categoryType][$category].Color
            }
        }
    }

    # Process prefix-based categories
    Write-ColorOutput "`n[Processing Prefix-Based Files]" "Yellow"
    foreach ($category in $Categories.PrefixBased.Keys) {
        $pattern = $Categories.PrefixBased[$category].Pattern
        $color = $Categories.PrefixBased[$category].Color
        $files = Get-ChildItem -Path $config.DesktopPath -Filter $pattern -File
        
        if ($files.Count -gt 0) {
            Write-ColorOutput "  ► Processing $category files..." $color
        }
        
        foreach ($file in $files) {
            if ($config.ShowProgress) {
                $processedItems++
                Write-Progress -Activity "Organizing Desktop" -Status "Processing $($file.Name)" `
                    -PercentComplete (($processedItems / $totalItems) * 100)
            }
            
            $destinationPath = Join-Path -Path $config.DesktopPath -ChildPath $category
            if (Move-ItemWithRetry -Source $file.FullName -Destination $destinationPath) {
                Write-ColorOutput "    ✓ Moved: $($file.Name)" $color
                $successfulMoves++
            } else {
                Write-ColorOutput "    ✗ Failed: $($file.Name)" "Red"
                $failedMoves++
            }
        }
    }

    # Process extension-based categories
    Write-ColorOutput "`n[Processing Extension-Based Files]" "Yellow"
    foreach ($category in $Categories.ExtensionBased.Keys) {
        if ($category -ne "Others") {
            $patterns = $Categories.ExtensionBased[$category].Pattern -split ','
            $color = $Categories.ExtensionBased[$category].Color
            
            foreach ($pattern in $patterns) {
                $files = Get-ChildItem -Path $config.DesktopPath -Filter $pattern.Trim() -File |
                    Where-Object {
                        $_.Directory.Name -eq (Split-Path -Path $config.DesktopPath -Leaf) -and
                        $_.Name -notmatch "^(TAEP_|BWS_|WWS_|CW_|PRP_)"
                    }
                
                foreach ($file in $files) {
                    if ($config.ShowProgress) {
                        $processedItems++
                        Write-Progress -Activity "Organizing Desktop" -Status "Processing $($file.Name)" `
                            -PercentComplete (($processedItems / $totalItems) * 100)
                    }
                    
                    $destinationPath = Join-Path -Path $config.DesktopPath -ChildPath $category
                    if (Move-ItemWithRetry -Source $file.FullName -Destination $destinationPath) {
                        Write-ColorOutput "    ✓ Moved: $($file.Name)" $color
                        $successfulMoves++
                    } else {
                        Write-ColorOutput "    ✗ Failed: $($file.Name)" "Red"
                        $failedMoves++
                    }
                }
            }
        }
    }

    # Process remaining files (Others category)
    $unmovedFiles = Get-UnmovedFiles -Path $config.DesktopPath
    if ($unmovedFiles.Count -gt 0) {
        Write-ColorOutput "`n[Processing Remaining Files]" "Yellow"
        $othersPath = Join-Path -Path $config.DesktopPath -ChildPath "Others"
        
        foreach ($file in $unmovedFiles) {
            if (Move-ItemWithRetry -Source $file.FullName -Destination $othersPath) {
                Write-ColorOutput "  ► Moved to Others: $($file.Name)" "Gray"
                $successfulMoves++
            } else {
                Write-ColorOutput "  ✗ Failed to move: $($file.Name)" "Red"
                $failedMoves++
                Add-Content -Path $config.UnmovedFilesReport -Value $file.FullName
            }
        }
    }

    # Create and move folders
    $FoldersPath = Join-Path -Path $config.DesktopPath -ChildPath "Folders"
    if (!(Test-Path -Path $FoldersPath -PathType Container)) {
        New-Item -Path $FoldersPath -ItemType Directory | Out-Null
        Write-ColorOutput "`n[Created Folders Directory]" "Green"
    }

    Get-ChildItem -Path $config.DesktopPath -Directory | 
        Where-Object {
            $_.FullName -ne $FoldersPath -and 
            $_.Name -notin $Categories.PrefixBased.Keys -and 
            $_.Name -notin $Categories.ExtensionBased.Keys
        } |
        ForEach-Object {
            if (Move-ItemWithRetry -Source $_.FullName -Destination $FoldersPath) {
                Write-ColorOutput "  ► Moved folder: $($_.Name)" "Blue"
            } else {
                Write-ColorOutput "  ✗ Failed to move folder: $($_.Name)" "Red"
            }
        }

    # Generate final report
    $finalUnmovedFiles = Get-UnmovedFiles -Path $config.DesktopPath
    if ($finalUnmovedFiles.Count -gt 0) {
        Write-ColorOutput "`n[Unmoved Files Report]" "Red"
        Write-ColorOutput "The following files could not be moved:" "Yellow"
        foreach ($file in $finalUnmovedFiles) {
            Write-ColorOutput "  • $($file.Name)" "Red"
            Add-Content -Path $config.UnmovedFilesReport -Value $file.FullName
        }
        Write-ColorOutput "Full report saved to: $($config.UnmovedFilesReport)" "Cyan"
    }

    # Display summary
    Write-ColorOutput "`n[Organization Summary]" "Cyan"
    Write-ColorOutput "  ✓ Total files processed: $totalItems" "Green"
    Write-ColorOutput "  ✓ Successfully moved: $successfulMoves" "Green"
    if ($failedMoves -gt 0) {
        Write-ColorOutput "  ✗ Failed to move: $failedMoves" "Red"
    }
    Write-ColorOutput "  ✓ Log file: $($config.LogPath)" "Green"
    if ($config.CreateBackup) {
        Write-ColorOutput "  ✓ Backup location: $backupPath" "Green"
    }
    Write-ColorOutput "`nDesktop organization completed!" "Green"
}
catch {
    Write-ToLog "Critical error occurred: $_" "ERROR"
    if ($config.CreateBackup) {
        Write-ColorOutput "`nBackup is available at: $backupPath" "Yellow"
    }
}

finally {
    if ($config.ShowProgress) {
        Write-Progress -Activity "Organizing Desktop" -Completed
    }
    Write-ColorOutput "`n=============================================" "Cyan"
    Write-ColorOutput "           Organization Complete" "Yellow"
    Write-ColorOutput "=============================================" "Cyan"
    
    # Display locations of important files
    Write-ColorOutput "`n[Important Locations]" "Magenta"
    Write-ColorOutput "• Log File: $($config.LogPath)" "White"
    if (Test-Path $config.UnmovedFilesReport) {
        Write-ColorOutput "• Unmoved Files Report: $($config.UnmovedFilesReport)" "White"
    }
    if ($backupPath) {
        Write-ColorOutput "• Backup Location: $backupPath" "White"
    }
    
    # Display organization tips
    Write-ColorOutput "`n[Maintenance Tips]" "Yellow"
    Write-ColorOutput "• Check the log file for detailed operation history" "Gray"
    Write-ColorOutput "• Review unmoved files report for any missed items" "Gray"
    Write-ColorOutput "• Backup is preserved for 7 days" "Gray"
    
    # Final newline for cleaner console output
    Write-ColorOutput "`n"
}
