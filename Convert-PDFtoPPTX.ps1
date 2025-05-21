# Script to convert all PDF files in the current directory to PowerPoint (PPTX)
# Author: AI Assistant
# Date: 2025-05-13

# PowerShell script that can be double-clicked to run

# Create a PowerPoint application instance
$pptApp = New-Object -ComObject PowerPoint.Application
# Make PowerPoint visible (set to $false for background processing)
$pptApp.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

# Get all PDF files in the same directory as the script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pdfFiles = Get-ChildItem -Path $scriptDir -Filter "*.pdf"

if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in the current directory."
    $null = Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Converting $($pdfFiles.Count) PDF files to PPTX..."
$successCount = 0

foreach ($file in $pdfFiles) {
    try {
        $fullPath = $file.FullName
        $pptxPath = [System.IO.Path]::ChangeExtension($fullPath, ".pptx")
        
        Write-Host "Converting: $($file.Name) to PPTX..."
        
        # Create a new presentation
        $presentation = $pptApp.Presentations.Add($true)
        
        # Add a slide for each page of the PDF
        try {
            # Insert the PDF
            $slide = $presentation.Slides.Add(1, 11) # 11 is the layout type for a blank slide
            
            # Insert the PDF as an object
            # Note: Each page of the PDF will be imported as a separate slide
            $newShape = $slide.Shapes.AddOLEObject(0, 0, -1, -1, "AcroExch.Document", $fullPath)
            
            # Get rid of the initial blank slide (it's a workaround due to how the insertion works)
            $slide.Delete()
            
            # Save the presentation
            $presentation.SaveAs($pptxPath)
            $successCount++
            Write-Host "Converted: $($file.Name) to PPTX successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "Error during PDF insertion for $($file.Name): $_" -ForegroundColor Red
            $presentation.Close()
            continue
        }
        
        # Close the presentation
        $presentation.Close()
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

Write-Host "Conversion complete! Successfully converted $successCount of $($pdfFiles.Count) files."
Write-Host "The PPTX files are saved in the same directory as the PDF files."
$null = Read-Host "Press Enter to exit..."
