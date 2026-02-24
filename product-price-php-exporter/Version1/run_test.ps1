# PowerShell test script for CSV to Excel converter
# This script automates testing by providing input to the converter

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "CSV to Excel Converter - Automated Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Test file path
$testFile = "h:\Repo\WordpressDevelopment\Products-Price-Exporter\test_sample.csv"

Write-Host "Test file: $testFile" -ForegroundColor Yellow
Write-Host ""

# Run the converter and provide input automatically
Write-Host "Running converter..." -ForegroundColor Green
$testFile | python "csv-to-excel-inpu.py"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Test completed!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
