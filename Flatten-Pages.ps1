# ============================================================
# Flatten-Pages.ps1
# "Prints to PDF" each PDF in a user-selected input folder,
# flattening form fields, annotations, and comments into the
# page content. Flattened copies are saved to the chosen
# output folder; originals are not modified.
# ============================================================

# --- Modern Folder Picker (with address bar) via Shell COM ---
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

function Select-Folder {
    param ([string]$Title)

    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.BrowseForFolder(0, $Title, 0x00000010 + 0x00000040, 0)
    #   0x00000010 = BIF_EDITBOX       (adds a text field to type/paste a path)
    #   0x00000040 = BIF_NEWDIALOGSTYLE (modern resizable dialog with address bar)

    if ($folder -and $folder.Self) {
        $path = $folder.Self.Path
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $path
    }

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
    return $null
}

# Prompt for input folder
$inputFolder = Select-Folder -Title "Select the INPUT folder containing PDF files"
if (-not $inputFolder) {
    Write-Host "No input folder selected. Exiting."
    exit
}
Write-Host "Input folder:  $inputFolder"

# Prompt for output folder
$outputFolder = Select-Folder -Title "Select the OUTPUT folder for flattened PDFs"
if (-not $outputFolder) {
    Write-Host "No output folder selected. Exiting."
    exit
}
Write-Host "Output folder: $outputFolder"

# Safety check: make sure output folder exists
if (-not (Test-Path -LiteralPath $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory | Out-Null
}

# --- Gather PDF files ---
$pdfFiles = Get-ChildItem -Path $inputFolder -Filter "*.pdf" -File

if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in the input folder. Exiting."
    exit
}

Write-Host "Found $($pdfFiles.Count) PDF file(s). Processing..."
Write-Host ""

# --- Counters and results list for summary ---
$totalProcessed = 0
$totalFlattened  = 0
$results         = @()

# --- Process each PDF ---
$acrobatApp = New-Object -ComObject AcroExch.App
$acrobatApp.Hide()

foreach ($pdf in $pdfFiles) {
    $filePath = $pdf.FullName
    $savePath = Join-Path $outputFolder $pdf.Name
    Write-Host "Processing: $($pdf.Name)"

    $avDoc = New-Object -ComObject AcroExch.AVDoc

    if (-not $avDoc.Open($filePath, "")) {
        Write-Host "  WARNING: Could not open '$($pdf.Name)'. Skipping."
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        continue
    }

    $pdDoc     = $avDoc.GetPDDoc()
    $pageCount = $pdDoc.GetNumPages()

    # Use the Acrobat JavaScript bridge to flatten all form fields,
    # annotations, and comments into the page content.
    $jsObj = $pdDoc.GetJSObject()
    $jsObj.flattenPages()

    # Save flattened PDF to the output folder (1 = PDSaveFull)
    $saveOk = $pdDoc.Save(1, $savePath)

    $avDoc.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null

    if ($saveOk) {
        Write-Host "  Flattened and saved ($pageCount page(s))."
        $totalProcessed++
        $totalFlattened++
        $results += [PSCustomObject]@{ FileName = $pdf.Name; Pages = $pageCount; Status = "Flattened" }
    } else {
        Write-Host "  ERROR: Save failed for '$($pdf.Name)'."
        $totalProcessed++
        $results += [PSCustomObject]@{ FileName = $pdf.Name; Pages = $pageCount; Status = "Error" }
    }
}

# --- Cleanup ---
$acrobatApp.Exit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "Done. All files processed to: $outputFolder"

# --- Export CSV to the output folder with a timestamp filename ---
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath   = Join-Path $outputFolder "$timestamp.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# --- Summary popup with option to view the CSV ---
$answer = [System.Windows.Forms.MessageBox]::Show(
    "Total Processed: $totalProcessed`nTotal Flattened: $totalFlattened`n`nWould you like to view the summary?",
    "Flatten Pages - Complete",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    Start-Process "excel.exe" -ArgumentList "`"$csvPath`""
}
