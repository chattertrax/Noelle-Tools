# ============================================================
# Remove-Pages.ps1
# Removes pages containing a specified phrase from all PDFs
# in a user-selected folder using Adobe Acrobat Pro COM objects.
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

# --- Gather PDF files ---
$pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -File

if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in the selected folder. Exiting."
    exit
}

Write-Host "Found $($pdfFiles.Count) PDF file(s). Processing..."
Write-Host ""

# --- Two-pass approach per PDF ---
# Pass 1: Open with AVDoc (read-only) for text extraction — the AV layer
#         is required for CreatePageHilite to work.  No modifications are
#         made, so closing triggers no save dialog even with GPO.
# Pass 2: Reopen with PDDoc (no UI), delete pages, save to a temp file,
#         close, and replace the original via PowerShell.

$acrobatApp = New-Object -ComObject AcroExch.App
$acrobatApp.Hide()

foreach ($pdf in $pdfFiles) {
    $filePath = $pdf.FullName
    Write-Host "Processing: $($pdf.Name)"

    # --- Pass 1: read-only text extraction via AVDoc ---
    $avDoc = New-Object -ComObject AcroExch.AVDoc

    if (-not $avDoc.Open($filePath, "")) {
        Write-Host "  WARNING: Could not open '$($pdf.Name)'. Skipping."
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        continue
    }

    $pdDoc = $avDoc.GetPDDoc()
    $pageCount = $pdDoc.GetNumPages()

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

    # Close AVDoc without any modifications — no save dialog
    $avDoc.Close($true)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null

    if ($pagesToRemove.Count -eq 0) {
        Write-Host "  No matching pages found. Skipping."
        continue
    }

    if ($pagesToRemove.Count -eq $pageCount) {
        Write-Host "  WARNING: All $pageCount page(s) match the phrase. Skipping file to avoid an empty document."
        continue
    }

    Write-Host "  Removing $($pagesToRemove.Count) of $pageCount page(s)..."

    # --- Pass 2: delete pages & save via PDDoc (no UI) ---
    $pdDoc2 = New-Object -ComObject AcroExch.PDDoc

    if (-not $pdDoc2.Open($filePath)) {
        Write-Host "  ERROR: Could not reopen '$($pdf.Name)' for editing."
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc2) | Out-Null
        continue
    }

    # Delete pages in reverse order so indices remain valid
    $pagesToRemove = $pagesToRemove | Sort-Object -Descending

    foreach ($pageIndex in $pagesToRemove) {
        $pdDoc2.DeletePages($pageIndex, $pageIndex)
    }

    # Save to temp file, close (releases lock), then replace original
    $tempPath = $filePath + ".tmp"
    $saveOk = $pdDoc2.Save(1, $tempPath)  # 1 = PDSaveFull

    $pdDoc2.Close() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc2) | Out-Null

    if ($saveOk) {
        Remove-Item -LiteralPath $filePath -Force
        Rename-Item -LiteralPath $tempPath -NewName (Split-Path $filePath -Leaf)
        $remainingPages = $pageCount - $pagesToRemove.Count
        Write-Host "  Saved. $remainingPages page(s) remaining."
    } else {
        Write-Host "  ERROR: Save failed for '$($pdf.Name)'."
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
    }
}

$acrobatApp.Exit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "Done. All files processed."
