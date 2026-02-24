# ==========================================
# PowerShell Test Script for CSV to Word
# ==========================================

Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " CSV to Word Converter Test Script          " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Define test files
$testCSV = Join-Path $scriptDir "test_sample.csv"
$productionCSV = Join-Path $scriptDir "vapeclub3-products-price.csv"

# Check if test file exists
if (-not (Test-Path $testCSV)) {
    Write-Host "[ERROR] Test file not found: $testCSV" -ForegroundColor Red
    Write-Host "Please make sure test_sample.csv exists in the same directory." -ForegroundColor Yellow
    exit 1
}

Write-Host "[INFO] Test file found: test_sample.csv" -ForegroundColor Green
Write-Host ""

# Ask user which test to run
Write-Host "Select test mode:" -ForegroundColor Yellow
Write-Host "  1 - Test with sample file (test_sample.csv - 9 rows)" -ForegroundColor White
Write-Host "  2 - Production test (vapeclub3-products-price.csv - 1533 rows)" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter choice 1 or 2"

if ($choice -eq "2") {
    if (-not (Test-Path $productionCSV)) {
        Write-Host "[ERROR] Production file not found: $productionCSV" -ForegroundColor Red
        exit 1
    }
    $inputFile = $productionCSV
    Write-Host "[INFO] Running production test..." -ForegroundColor Cyan
} else {
    $inputFile = $testCSV
    Write-Host "[INFO] Running test with sample file..." -ForegroundColor Cyan
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Starting conversion...                      " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Run the Python script with automatic input
$inputFile | python csv-to-word.py

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host " Test completed!                             " -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Check if output file was created
$outputFile = $inputFile -replace "\.csv$", ".docx"

if (Test-Path $outputFile) {
    Write-Host "[SUCCESS] Output file created successfully!" -ForegroundColor Green
    Write-Host "Location: $outputFile" -ForegroundColor White
    Write-Host ""
    
    # Get file info
    $fileInfo = Get-Item $outputFile
    $fileSizeKB = [math]::Round($fileInfo.Length/1024, 2)
    
    Write-Host "File Details:" -ForegroundColor Yellow
    Write-Host "  - Size: $($fileInfo.Length) bytes - $fileSizeKB KB" -ForegroundColor White
    Write-Host "  - Created: $($fileInfo.CreationTime)" -ForegroundColor White
    Write-Host "  - Modified: $($fileInfo.LastWriteTime)" -ForegroundColor White
    Write-Host ""
    
    # Ask if user wants to open the file
    $openFile = Read-Host "Do you want to open the Word file? Y/N"
    if ($openFile -eq "Y" -or $openFile -eq "y") {
        Write-Host "[INFO] Opening Word file..." -ForegroundColor Cyan
        Start-Process $outputFile
    }
} else {
    Write-Host "[ERROR] Output file was not created!" -ForegroundColor Red
    Write-Host "Please check the error messages above." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
