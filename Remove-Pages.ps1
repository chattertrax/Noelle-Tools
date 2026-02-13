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

# --- Process each PDF using PDDoc only (no AVDoc UI layer) ---
# Working entirely at the PD layer avoids all UI dialogs, including
# GPO-enforced save confirmations that cannot be suppressed.

foreach ($pdf in $pdfFiles) {
    $filePath = $pdf.FullName
    Write-Host "Processing: $($pdf.Name)"

    $pdDoc = New-Object -ComObject AcroExch.PDDoc

    if (-not $pdDoc.Open($filePath)) {
        Write-Host "  WARNING: Could not open '$($pdf.Name)'. Skipping."
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        continue
    }

    $pageCount = $pdDoc.GetNumPages()

    # Collect indices of pages to remove.
    # Uses native Acrobat COM text-selection objects (HiliteList,
    # PDPage, PDTextSelect) — all PD-layer, no UI involved.
    $pagesToRemove = @()

    for ($i = 0; $i -lt $pageCount; $i++) {
        $pdPage   = $pdDoc.AcquirePage($i)
        $hilite   = New-Object -ComObject AcroExch.HiliteList
        $hilite.Add(0, 32767)                         # select all text on the page
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
        $pdDoc.Close() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        continue
    }

    if ($pagesToRemove.Count -eq $pageCount) {
        Write-Host "  WARNING: All $pageCount page(s) match the phrase. Skipping file to avoid an empty document."
        $pdDoc.Close() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        continue
    }

    Write-Host "  Removing $($pagesToRemove.Count) of $pageCount page(s)..."

    # Delete pages in reverse order so indices remain valid
    $pagesToRemove = $pagesToRemove | Sort-Object -Descending

    foreach ($pageIndex in $pagesToRemove) {
        $pdDoc.DeletePages($pageIndex, $pageIndex)
    }

    # Save over the original file
    $pdDoc.Save(1, $filePath)  # 1 = PDSaveFull

    $pdDoc.Close() | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null

    $remainingPages = $pageCount - $pagesToRemove.Count
    Write-Host "  Saved. $remainingPages page(s) remaining."
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "Done. All files processed."
