# Script to convert all PowerPoint (PPTX) and PDF files in the current directory to PDF
# Author: AI Assistant
# Date: 2025-05-13

# PowerShell script that can be double-clicked to run

# Create a PowerPoint application instance
$pptApp = New-Object -ComObject PowerPoint.Application
# Make PowerPoint visible (set to $false for background processing)
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Get all PPTX and PDF files in the same directory as the script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pptxFiles = Get-ChildItem -Path $scriptDir -Filter "*.pptx"
$pdfFiles = Get-ChildItem -Path $scriptDir -Filter "*.pdf"

$totalFiles = $pptxFiles.Count + $pdfFiles.Count

if ($totalFiles -eq 0) {
    Write-Host "No PPTX or PDF files found in the current directory."
    $null = Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Found $($pptxFiles.Count) PPTX files and $($pdfFiles.Count) PDF files."
Write-Host "Converting all files to PDF format..."

$successCount = 0

# Function to ensure unique filename
function Get-UniqueFileName {
    param (
        [string]$FilePath
    )
    
    $folder = [System.IO.Path]::GetDirectoryName($FilePath)
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $extension = [System.IO.Path]::GetExtension($FilePath)
    
    $counter = 1
    $newPath = $FilePath
    
    while (Test-Path $newPath) {
        $newPath = Join-Path -Path $folder -ChildPath "$filename`_$counter$extension"
        $counter++
    }
    
    return $newPath
}

# Convert PPTX files to PDF
foreach ($file in $pptxFiles) {
    try {
        $fullPath = $file.FullName
        $pdfPath = [System.IO.Path]::ChangeExtension($fullPath, ".pdf")
        
        # If the PDF already exists, create a unique filename
        if (Test-Path $pdfPath) {
            $pdfPath = Get-UniqueFileName -FilePath $pdfPath
        }
        
        Write-Host "Converting PPTX: $($file.Name) to PDF..."
        
        # Open the presentation
        $presentation = $pptApp.Presentations.Open($fullPath)
        
        # Save as PDF
        $presentation.SaveAs($pdfPath, 32)  # 32 is the constant for PDF format
        
        # Close the presentation
        $presentation.Close()
        
        $successCount++
        Write-Host "Converted: $($file.Name) to PDF successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error converting $($file.Name): $_" -ForegroundColor Red
    }
}

# Process existing PDF files (just copy or rename them if needed)
foreach ($file in $pdfFiles) {
    try {
        $fullPath = $file.FullName
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
        $directory = [System.IO.Path]::GetDirectoryName($fullPath)
        $newPdfPath = Join-Path -Path $directory -ChildPath "$baseName`_converted.pdf"
        
        # If the new PDF already exists, create a unique filename
        if (Test-Path $newPdfPath) {
            $newPdfPath = Get-UniqueFileName -FilePath $newPdfPath
        }
        
        Write-Host "Processing PDF: $($file.Name)..."
        
        # Simply copy the PDF with a new name to include it in the output
        Copy-Item -Path $fullPath -Destination $newPdfPath
        
        $successCount++
        Write-Host "Processed: $($file.Name) successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error processing $($file.Name): $_" -ForegroundColor Red
    }
}

# Quit PowerPoint
$pptApp.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Conversion complete! Successfully processed $successCount of $totalFiles files."
Write-Host "All files have been converted to PDF format and saved in the script directory."
$null = Read-Host "Press Enter to exit..."
