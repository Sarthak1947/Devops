# ============================================
 
# Script: MapAndDownload.ps1
 
# Description:
#   - Maps a network drive.
#   - Downloads the Invoke-FslShrinkDisk script from GitHub.
#   - Executes the Invoke-FslShrinkDisk script to shrink disk usage on the mapped drive.
#   - Converts the Shrink_Result.csv to ShrinkResult.xlsx after completion.
#   - Unmaps the network drive after execution, even if the script is stopped or interrupted.
 
# ============================================
 
# ---------------------------
# Configuration Variables
# ---------------------------
 
# Network Drive Mapping Parameters
$DriveLetter = "U:"  # Updated drive letter
$NetworkPath = "\\DC10SMBANP-02a2.molina.mhc\sc-anf-avd-mcw-prod-01"  # Updated network path
 
# GitHub Repository Details
$gitUrl = "https://github.com/FSLogix/Invoke-FslShrinkDisk.git"
 
# Base path for downloading the repository
$basePath = "C:\Login_Script\MCW"  # Updated base path
 
# Destination path for the repository
$destinationPath = Join-Path -Path $basePath -ChildPath "Invoke-FslShrinkDisk"  # Updated download path
 
# Path to the Shrink Script after cloning
$ShrinkScriptPath = Join-Path -Path $destinationPath -ChildPath "Invoke-FslShrinkDisk.ps1"
 
# Path for the Shrink Log File
$ShrinkLogFilePath = "U:\Shrink_Result.csv"  # Updated drive letter for log file
 
# Path for the Excel Result File
$ShrinkExcelPath = Join-Path -Path $destinationPath -ChildPath "ShrinkResult.xlsx"  # Excel file path
 
# ---------------------------
# Function: Map Network Drive
# ---------------------------
 
function Map-NetworkDrive {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter,
 
        [Parameter(Mandatory = $true)]
        [string]$NetworkPath
    )
 
    try {
        # Extract the drive name without the colon
        $DriveName = $DriveLetter.TrimEnd(':')
 
        # Check if the drive is already mapped
        $mappedDrive = Get-PSDrive -Name $DriveName -ErrorAction SilentlyContinue
 
        if ($mappedDrive) {
            # If the drive is mapped to the desired path, do nothing
            if ($mappedDrive.DisplayRoot -eq $NetworkPath) {
                Write-Host "Drive $DriveLetter is already mapped to $NetworkPath." -ForegroundColor Green
            }
            else {
                Write-Warning "Drive $DriveLetter is mapped to a different path. Remapping to $NetworkPath."
                # Remove the existing mapping
                net use $DriveLetter /delete /y | Out-Null
                Start-Sleep -Seconds 2
                # Map to the new network path
                net use $DriveLetter $NetworkPath /persistent:no
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "Drive $DriveLetter successfully remapped to $NetworkPath." -ForegroundColor Green
                }
                else {
                    Write-Error "Failed to remap drive $DriveLetter to $NetworkPath."
                    exit 1
                }
            }
        }
        else {
            # Map the drive
            Write-Host "Mapping drive $DriveLetter to $NetworkPath." -ForegroundColor Cyan
            net use $DriveLetter $NetworkPath /persistent:no
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Drive $DriveLetter successfully mapped to $NetworkPath." -ForegroundColor Green
            }
            else {
                Write-Error "Failed to map drive $DriveLetter to $NetworkPath."
                exit 1
            }
        }
    }
    catch {
        Write-Error "An error occurred while mapping the network drive: $_"
        exit 1
    }
}
 
# ---------------------------
# Function: Unmap Network Drive
# ---------------------------
 
function Unmap-NetworkDrive {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter
    )
 
    try {
        Write-Host "Unmapping drive $DriveLetter." -ForegroundColor Cyan
        net use $DriveLetter /delete /y | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Drive $DriveLetter successfully unmapped." -ForegroundColor Green
        }
        else {
            Write-Warning "Failed to unmap drive $DriveLetter."
        }
    }
    catch {
        Write-Error "An error occurred while unmapping the network drive: $_"
    }
}
 
# ---------------------------
# Function: Download GitHub Repository
# ---------------------------
 
function Download-GitHubRepository {
    param (
        [Parameter(Mandatory = $true)]
        [string]$gitUrl,
 
        [Parameter(Mandatory = $true)]
        [string]$destinationPath
    )
 
    try {
        Write-Host "Downloading GitHub repository from: $gitUrl" -ForegroundColor Cyan
 
        # Check if destination directory exists
        if (Test-Path -Path $destinationPath) {
            # Check if it's a Git repository
            $gitFolder = Join-Path -Path $destinationPath -ChildPath ".git"
            if (Test-Path -Path $gitFolder) {
                Write-Host "Destination directory is a Git repository. Pulling latest changes..." -ForegroundColor Yellow
                git -C $destinationPath pull
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "Repository successfully updated at: $destinationPath" -ForegroundColor Green
                }
                else {
                    Write-Error "Failed to pull updates from the repository."
                    exit 1
                }
            }
            else {
                Write-Warning "Destination directory exists but is not a Git repository. Removing and cloning afresh." -ForegroundColor Yellow
                Remove-Item -Path $destinationPath -Recurse -Force
                Start-Sleep -Seconds 2
                git clone $gitUrl $destinationPath
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "Repository successfully cloned to: $destinationPath" -ForegroundColor Green
                }
                else {
                    Write-Error "Failed to clone the repository from $gitUrl."
                    exit 1
                }
            }
        }
        else {
            # Clone the repository using Git
            git clone $gitUrl $destinationPath
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Repository successfully cloned to: $destinationPath" -ForegroundColor Green
            }
            else {
                Write-Error "Failed to clone the repository from $gitUrl."
                exit 1
            }
        }
    }
    catch {
        Write-Error "Failed to download the repository. Error details: $_"
        exit 1
    }
}
 
# ---------------------------
# Function: Convert CSV to Excel
# ---------------------------
 
function Convert-CsvToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$csvPath,
 
        [Parameter(Mandatory = $true)]
        [string]$excelPath
    )
 
    try {
        Write-Host "Converting CSV to Excel: $excelPath" -ForegroundColor Cyan
 
        # Create Excel COM object
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
 
        # Open the CSV file
        $workbook = $excel.Workbooks.Open($csvPath)
 
        # Save as Excel workbook (xlOpenXMLWorkbook format = 51)
        $workbook.SaveAs($excelPath, 51)
 
        # Close the workbook and quit Excel
        $workbook.Close()
        $excel.Quit()
 
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
 
        Write-Host "Excel file created at: $excelPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to convert CSV to Excel: $_"
        exit 1
    }
}
 
# ---------------------------
# Function: Cleanup on Exit
# ---------------------------
 
function Cleanup {
    Unmap-NetworkDrive -DriveLetter $DriveLetter
}
 
# Register the Cleanup function to run when the script is exiting
Register-EngineEvent PowerShell.Exiting -Action { Cleanup } | Out-Null
 
# ---------------------------
# Main Script Execution
# ---------------------------
 
# Wrap main execution in a try/finally to ensure Cleanup runs
try {
    # Step 1: Map the Network Drive
    Map-NetworkDrive -DriveLetter $DriveLetter -NetworkPath $NetworkPath
 
    # Step 2: Download the GitHub Repository
    Download-GitHubRepository -gitUrl $gitUrl -destinationPath $destinationPath
 
    # Step 3: Execute the Shrink Script
    Write-Host "Executing Invoke-FslShrinkDisk script..." -ForegroundColor Cyan
    & $ShrinkScriptPath -Path $DriveLetter -Recurse -LogFilePath $ShrinkLogFilePath
 
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Shrink script executed successfully." -ForegroundColor Green
 
        # Step 4: Convert Shrink_Result.csv to ShrinkResult.xlsx
        Convert-CsvToExcel -csvPath $ShrinkLogFilePath -excelPath $ShrinkExcelPath
    }
    else {
        Write-Error "Shrink script encountered errors during execution."
        exit 1
    }
}
finally {
    # Cleanup: Unmap the network drive
    Cleanup
}
 
# Optional: Open the Shrink Log File (Uncomment if desired)
# Start-Process "notepad.exe" $ShrinkLogFilePath
 
# End of Script
