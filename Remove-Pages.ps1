# ============================================================
# Remove-Pages.ps1
# Removes pages containing a specified phrase from all PDFs
# in a user-selected input folder. All PDFs are copied to the
# chosen output folder -- matching pages are removed from the
# modified copies; unmatched files are copied unchanged.
# ============================================================

# --- Configuration ---
# Set the phrase to search for (case-insensitive match).
# Any page containing this phrase anywhere in its text will be removed.
$Phrase = "Information Missing on the Document"

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
$outputFolder = Select-Folder -Title "Select the OUTPUT folder for processed PDFs"
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

# --- Counters for summary ---
$totalProcessed = 0
$totalModified  = 0

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

    $pdDoc = $avDoc.GetPDDoc()
    $pageCount = $pdDoc.GetNumPages()

    # --- Text extraction: find pages that contain the phrase ---
    $pagesToRemove = @()

    for ($i = 0; $i -lt $pageCount; $i++) {
        $pdPage   = $pdDoc.AcquirePage($i)
        $hilite   = New-Object -ComObject AcroExch.HiliteList
        $hilite.Add(0, 32767)
        $textSelect = $pdPage.CreatePageHilite($hilite)

        $pageText = ""
        if ($textSelect) {
            $numText = $textSelect.GetNumText()
            for ($t = 0; $t -lt $numText; $t++) {
                $pageText += $textSelect.GetText($t)
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($textSelect) | Out-Null
        }

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($hilite)  | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdPage)  | Out-Null

        if ($pageText -match [regex]::Escape($Phrase)) {
            $pagesToRemove += $i
        }
    }

    # --- No matching pages: save an unmodified copy ---
    if ($pagesToRemove.Count -eq 0) {
        Write-Host "  No matching pages. Copying unchanged."
        $pdDoc.Save(1, $savePath) | Out-Null
        $avDoc.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        $totalProcessed++
        continue
    }

    # --- All pages match: skip to avoid an empty document ---
    if ($pagesToRemove.Count -eq $pageCount) {
        Write-Host "  WARNING: All $pageCount page(s) match the phrase. Skipping to avoid an empty document."
        $avDoc.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        continue
    }

    Write-Host "  Removing $($pagesToRemove.Count) of $pageCount page(s)..."

    # --- Delete pages in reverse order so indices stay valid ---
    $pagesToRemove = $pagesToRemove | Sort-Object -Descending

    foreach ($pageIndex in $pagesToRemove) {
        $pdDoc.DeletePages($pageIndex, $pageIndex)
    }

    # --- Save modified PDF to the output folder ---
    $saveOk = $pdDoc.Save(1, $savePath)  # 1 = PDSaveFull

    # Save clears the dirty flag, so Close won't trigger a GPO dialog
    $avDoc.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null

    if ($saveOk) {
        $remainingPages = $pageCount - $pagesToRemove.Count
        Write-Host "  Saved ($remainingPages page(s) remaining)."
        $totalProcessed++
        $totalModified++
    } else {
        Write-Host "  ERROR: Save failed for '$($pdf.Name)'."
    }
}

# --- Cleanup ---
$acrobatApp.Exit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "Done. All files processed to: $outputFolder"

# --- Summary popup ---
[System.Windows.Forms.MessageBox]::Show(
    "Total Processed: $totalProcessed`nTotal Modified: $totalModified",
    "Remove Pages - Complete",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null
