# ============================================================
# Remove-Pages.ps1
# Removes pages containing a specified phrase from all PDFs
# in a user-selected folder using Adobe Acrobat Pro COM objects.
# Modified PDFs are saved to an "Output" subfolder; originals
# are left untouched.
# ============================================================

# --- Configuration ---
# Set the phrase to search for (case-insensitive match).
# Any page containing this phrase anywhere in its text will be removed.
$Phrase = "Information Missing on the Document"

# --- Folder Selection Dialog ---
Add-Type -AssemblyName System.Windows.Forms

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder containing PDF files"
$folderBrowser.ShowNewFolderButton = $false

$result = $folderBrowser.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "No folder selected. Exiting."
    exit
}

$folderPath = $folderBrowser.SelectedPath
Write-Host "Selected folder: $folderPath"

# --- Create Output subfolder ---
$outputPath = Join-Path $folderPath "Output"
if (-not (Test-Path -LiteralPath $outputPath)) {
    New-Item -Path $outputPath -ItemType Directory | Out-Null
}

# --- Gather PDF files ---
$pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -File

if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in the selected folder. Exiting."
    exit
}

Write-Host "Found $($pdfFiles.Count) PDF file(s). Processing..."
Write-Host ""

# --- Process each PDF ---
$acrobatApp = New-Object -ComObject AcroExch.App
$acrobatApp.Hide()

foreach ($pdf in $pdfFiles) {
    $filePath = $pdf.FullName
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

    if ($pagesToRemove.Count -eq 0) {
        Write-Host "  No matching pages found. Skipping."
        $avDoc.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        continue
    }

    if ($pagesToRemove.Count -eq $pageCount) {
        Write-Host "  WARNING: All $pageCount page(s) match the phrase. Skipping file to avoid an empty document."
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

    # --- Save modified PDF to the Output subfolder ---
    $savePath = Join-Path $outputPath $pdf.Name
    $saveOk  = $pdDoc.Save(1, $savePath)  # 1 = PDSaveFull

    # After Save the dirty flag is cleared, so Close won't prompt
    $avDoc.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null

    if ($saveOk) {
        $remainingPages = $pageCount - $pagesToRemove.Count
        Write-Host "  Saved to Output\$($pdf.Name) — $remainingPages page(s) remaining."
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
Write-Host "Done. All files processed. Modified PDFs are in: $outputPath"
