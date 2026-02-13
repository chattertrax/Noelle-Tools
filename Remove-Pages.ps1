# ============================================================
# Remove-Pages.ps1
# Removes pages containing a specified phrase from all PDFs
# in a user-selected folder using Adobe Acrobat Pro COM objects.
# ============================================================

# --- Configuration ---
# Set the phrase to search for (case-insensitive match).
# Any page containing this phrase anywhere in its text will be removed.
$Phrase = "REPLACE WITH YOUR PHRASE"

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

# --- Initialize Acrobat COM object ---
$acrobatApp = New-Object -ComObject AcroExch.App

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

    # Collect indices of pages to remove
    $pagesToRemove = @()
    $jsObj = $pdDoc.GetJSObject()

    for ($i = 0; $i -lt $pageCount; $i++) {
        # Extract full page text via Acrobat JavaScript bridge
        $numWords = $jsObj.getPageNumWords($i)
        $pageTextBuilder = New-Object System.Text.StringBuilder

        for ($w = 0; $w -lt $numWords; $w++) {
            $word = $jsObj.getPageNthWord($i, $w, $false)
            if ($w -gt 0) {
                [void]$pageTextBuilder.Append(" ")
            }
            [void]$pageTextBuilder.Append($word)
        }

        $pageText = $pageTextBuilder.ToString()

        if ($pageText -match [regex]::Escape($Phrase)) {
            $pagesToRemove += $i
        }
    }

    if ($pagesToRemove.Count -eq 0) {
        Write-Host "  No matching pages found. Skipping."
        $avDoc.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
        continue
    }

    if ($pagesToRemove.Count -eq $pageCount) {
        Write-Host "  WARNING: All $pageCount page(s) match the phrase. Skipping file to avoid an empty document."
        $avDoc.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
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

    $avDoc.Close($false)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null

    $remainingPages = $pageCount - $pagesToRemove.Count
    Write-Host "  Saved. $remainingPages page(s) remaining."
}

# --- Cleanup ---
$acrobatApp.Exit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "Done. All files processed."
