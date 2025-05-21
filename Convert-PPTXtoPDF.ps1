# Script to convert all PowerPoint (PPTX) files in the current directory to PDF
# Author: AI Assistant
# Date: 2025-05-13

# PowerShell script that can be double-clicked to run

# Create a PowerPoint application instance
$pptApp = New-Object -ComObject PowerPoint.Application
# Make PowerPoint visible (set to $false for background processing)
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Get all PPTX files in the same directory as the script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pptxFiles = Get-ChildItem -Path $scriptDir -Filter "*.pptx"

if ($pptxFiles.Count -eq 0) {
    Write-Host "No PPTX files found in the current directory."
    $null = Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Converting $($pptxFiles.Count) PPTX files to PDF..."
$successCount = 0

foreach ($file in $pptxFiles) {
    try {
        $fullPath = $file.FullName
        $pdfPath = [System.IO.Path]::ChangeExtension($fullPath, ".pdf")
        
        Write-Host "Converting: $($file.Name) to PDF..."
        
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

# Quit PowerPoint
$pptApp.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Conversion complete! Successfully converted $successCount of $($pptxFiles.Count) files."
Write-Host "The PDF files are saved in the same directory as the PPTX files."
$null = Read-Host "Press Enter to exit..."
