# Get Windows Updates
function Get-WindowsUpdates {
    try {
        $Session = New-Object -ComObject "Microsoft.Update.Session"
        $Searcher = $Session.CreateUpdateSearcher()
        $Result = $Searcher.Search("IsInstalled=0 and Type='Software'")
        
        if ($Result.Updates.Count -gt 0) {
            Write-Host "`nPending Windows Updates:" -ForegroundColor Yellow
            $Result.Updates | ForEach-Object {
                Write-Host "- $($_.Title)" -ForegroundColor Cyan
            }
        } else {
            Write-Host "No pending Windows updates." -ForegroundColor Green
        }
    } catch {
        Write-Host "Error checking Windows Updates: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Get Software from Package Managers
function Get-PackageUpdates {
    # Check Chocolatey
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Host "`nChecking Chocolatey packages..." -ForegroundColor Yellow
        $outdated = choco outdated
        if ($outdated -match "Chocolatey has determined") {
            Write-Host $outdated -ForegroundColor Cyan
        } else {
            Write-Host "All Chocolatey packages are up to date." -ForegroundColor Green
        }
    }

    # Check Winget
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "`nChecking Winget packages..." -ForegroundColor Yellow
        $wingetUpdates = winget upgrade
        if ($wingetUpdates -match "Available") {
            Write-Host $wingetUpdates -ForegroundColor Cyan
        } else {
            Write-Host "All Winget packages are up to date." -ForegroundColor Green
        }
    }
}

# Main execution
Write-Host "Checking for software updates..." -ForegroundColor Green
Get-WindowsUpdates
Get-PackageUpdates