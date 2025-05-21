#Requires -Version 5.0
<#
.SYNOPSIS
    Merges multiple PDF files into a single PDF document.

.DESCRIPTION
    This script uses the pdfunite.exe tool from MiKTeX to merge multiple PDF files into a single PDF document.
    The user can select individual PDF files or a folder containing PDF files to merge.

.PARAMETER InputFolder
    Specifies the folder containing PDF files to merge. If this parameter is not specified, 
    the user will be prompted to select PDF files individually.

.PARAMETER OutputFile
    Specifies the path for the merged PDF document. If this parameter is not specified,
    the user will be prompted to specify an output file.

.EXAMPLE
    .\Merge-PDFFiles.ps1
    Prompts the user to select PDF files and specify an output file.

.EXAMPLE
    .\Merge-PDFFiles.ps1 -InputFolder "C:\PDFs" -OutputFile "C:\Merged\merged.pdf"
    Merges all PDF files in the C:\PDFs folder into C:\Merged\merged.pdf.

.NOTES
    Author: AI Assistant
    Date: 2025-05-14
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Folder containing PDF files to merge")]
    [string]$InputFolder,

    [Parameter(Mandatory = $false, HelpMessage = "Output file path for the merged PDF")]
    [string]$OutputFile
)

# Function to test if pdfunite.exe is available
function Test-PDFUniteAvailable {
    $pdfunitePath = "C:\Users\bebag\AppData\Local\Programs\MiKTeX\miktex\bin\x64\pdfunite.exe"
    
    if (Test-Path -Path $pdfunitePath) {
        return $pdfunitePath
    }
    else {
        Write-Error "pdfunite.exe not found at: $pdfunitePath"
        Write-Error "Please make sure MiKTeX is installed correctly."
        return $false
    }
}

# Function to select PDF files using GUI
function Select-PDFFiles {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect = $true
        Filter = "PDF Files (*.pdf)|*.pdf"
        Title = "Select PDF Files to Merge"
    }

    if ($FileBrowser.ShowDialog() -eq 'OK') {
        return $FileBrowser.FileNames
    }
    else {
        Write-Host "No files selected. Exiting."
        return $null
    }
}

# Function to select output file using GUI
function Select-OutputFile {
    $FileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Filter = "PDF Files (*.pdf)|*.pdf"
        Title = "Save Merged PDF As"
        DefaultExt = "pdf"
        AddExtension = $true
    }

    if ($FileBrowser.ShowDialog() -eq 'OK') {
        return $FileBrowser.FileName
    }
    else {
        Write-Host "No output file selected. Exiting."
        return $null
    }
}

# Function to validate PDF files
function Test-PDFFiles {
    param (
        [string[]]$FilePaths
    )

    $validFiles = @()
    $invalidFiles = @()

    foreach ($file in $FilePaths) {
        if (Test-Path -Path $file -PathType Leaf) {
            $extension = [System.IO.Path]::GetExtension($file)
            if ($extension -eq ".pdf") {
                try {
                    # Try to open the file to make sure it's not locked
                    $stream = [System.IO.File]::Open($file, 'Open', 'Read', 'ReadWrite')
                    $stream.Close()
                    $stream.Dispose()
                    $validFiles += $file
                }
                catch {
                    Write-Warning "File is locked or unreadable: $file"
                    $invalidFiles += $file
                }
            }
            else {
                Write-Warning "Not a PDF file: $file"
                $invalidFiles += $file
            }
        }
        else {
            Write-Warning "File does not exist: $file"
            $invalidFiles += $file
        }
    }

    return @{
        ValidFiles = $validFiles
        InvalidFiles = $invalidFiles
    }
}

# Function to merge PDF files
function Merge-PDFs {
    param (
        [string[]]$InputFiles,
        [string]$OutputFilePath,
        [string]$PDFUnitePath
    )

    try {
        # Build the arguments for pdfunite
        $arguments = @()
        foreach ($file in $InputFiles) {
            $arguments += "`"$file`""
        }
        $arguments += "`"$OutputFilePath`""

        $argumentsString = $arguments -join " "
        
        # Execute pdfunite
        Write-Host "Merging PDF files... Please wait."
        $process = Start-Process -FilePath $PDFUnitePath -ArgumentList $argumentsString -NoNewWindow -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "PDF files merged successfully into: $OutputFilePath" -ForegroundColor Green
            return $true
        }
        else {
            Write-Error "Error merging PDF files. Exit code: $($process.ExitCode)"
            return $false
        }
    }
    catch {
        Write-Error "An error occurred while merging PDF files: $_"
        return $false
    }
}

# Main script execution starts here
# Load Windows Forms for file dialogs
Add-Type -AssemblyName System.Windows.Forms

# Check if pdfunite is available
$pdfunitePath = Test-PDFUniteAvailable
if (-not $pdfunitePath) {
    exit 1
}

# Get PDF files to merge
$pdfFiles = @()

if ($InputFolder) {
    # Use specified input folder
    if (Test-Path -Path $InputFolder -PathType Container) {
        $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf" | Select-Object -ExpandProperty FullName
        if ($pdfFiles.Count -eq 0) {
            Write-Warning "No PDF files found in folder: $InputFolder"
            exit 1
        }
    }
    else {
        Write-Error "Specified input folder does not exist: $InputFolder"
        exit 1
    }
}
else {
    # Prompt user to select files
    $pdfFiles = Select-PDFFiles
    if (-not $pdfFiles) {
        exit 1
    }
}

# Validate PDF files
$validationResult = Test-PDFFiles -FilePaths $pdfFiles
$validFiles = $validationResult.ValidFiles
$invalidFiles = $validationResult.InvalidFiles

if ($validFiles.Count -lt 2) {
    Write-Error "At least two valid PDF files are required for merging."
    if ($invalidFiles.Count -gt 0) {
        Write-Warning "Found $($invalidFiles.Count) invalid files:"
        $invalidFiles | ForEach-Object { Write-Warning " - $_" }
    }
    exit 1
}

# Get output file path
if (-not $OutputFile) {
    $OutputFile = Select-OutputFile
    if (-not $OutputFile) {
        exit 1
    }
}
else {
    # Ensure the directory exists
    $outputDir = [System.IO.Path]::GetDirectoryName($OutputFile)
    if (-not (Test-Path -Path $outputDir -PathType Container)) {
        Write-Error "Output directory does not exist: $outputDir"
        exit 1
    }
    
    # Ensure the output file has .pdf extension
    if ([System.IO.Path]::GetExtension($OutputFile) -ne ".pdf") {
        $OutputFile = [System.IO.Path]::ChangeExtension($OutputFile, ".pdf")
    }
}

# Display summary before merging
Write-Host "`nMerging PDF Files:" -ForegroundColor Cyan
$validFiles | ForEach-Object { Write-Host " - $([System.IO.Path]::GetFileName($_))" }
Write-Host "Output file will be saved to: $OutputFile" -ForegroundColor Cyan
Write-Host "Total files to merge: $($validFiles.Count)" -ForegroundColor Cyan

# Confirm with user
$confirmation = Read-Host "`nDo you want to proceed with merging these files? (Y/N)"
if ($confirmation.ToUpper() -ne "Y") {
    Write-Host "Operation canceled by user."
    exit 0
}

# Merge PDF files
$mergeResult = Merge-PDFs -InputFiles $validFiles -OutputFilePath $OutputFile -PDFUnitePath $pdfunitePath

if ($mergeResult) {
    if (Test-Path -Path $OutputFile) {
        Write-Host "`nSuccessfully created merged PDF file: $OutputFile" -ForegroundColor Green
        
        # Ask if user wants to open the file
        $openFile = Read-Host "Do you want to open the merged PDF file now? (Y/N)"
        if ($openFile.ToUpper() -eq "Y") {
            Invoke-Item $OutputFile
        }
    }
    else {
        Write-Error "Merge operation reported success, but the output file does not exist."
    }
}

