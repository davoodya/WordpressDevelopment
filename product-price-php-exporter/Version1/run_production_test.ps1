# PowerShell production test script
# Tests with actual vapeclub3-products-price.csv file

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "CSV to Excel Converter - PRODUCTION TEST" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Production file path
$productionFile = "h:\Repo\WordpressDevelopment\Products-Price-Exporter\vapeclub3-products-price.csv"

Write-Host "Production file: $productionFile" -ForegroundColor Yellow
Write-Host ""

# Check if file exists
if (-Not (Test-Path $productionFile)) {
    Write-Host "ERROR: File not found!" -ForegroundColor Red
    exit 1
}

# Count rows in CSV
$rowCount = (Get-Content $productionFile | Measure-Object -Line).Lines
Write-Host "CSV file has $rowCount rows" -ForegroundColor Green
Write-Host ""

# Run the converter
Write-Host "Running converter..." -ForegroundColor Green
Write-Host ""
$productionFile | python "csv-to-excel-inpu.py"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Production test completed!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Check output file: vapeclub3-products-price.xlsx" -ForegroundColor Yellow
Write-Host "Check log file: conversion_*.log" -ForegroundColor Yellow
