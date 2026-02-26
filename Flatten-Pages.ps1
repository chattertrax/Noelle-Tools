# ============================================================
# Flatten-Pages.ps1
# "Prints to PDF" each PDF in a user-selected input folder,
# producing a flat copy with no form fields, annotations, or
# interactive elements. Originals are not modified.
#
# Uses the "Adobe PDF" virtual printer (included with Acrobat
# Pro) configured for silent output to the chosen output folder.
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

# --- Check for "Adobe PDF" printer (ships with Acrobat Pro) ---
$printerName = "Adobe PDF"
$printerCheck = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
if (-not $printerCheck) {
    [System.Windows.Forms.MessageBox]::Show(
        "The 'Adobe PDF' printer was not found.`n`nThis printer is installed with Adobe Acrobat Pro and is required for flattening PDFs.",
        "Flatten Pages - Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit
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

# Safety: input and output must differ to avoid overwriting originals
if ($inputFolder.TrimEnd('\') -eq $outputFolder.TrimEnd('\')) {
    [System.Windows.Forms.MessageBox]::Show(
        "Input and output folders must be different to protect your original files.",
        "Flatten Pages - Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    exit
}

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

# --- Configure "Adobe PDF" printer for silent output ---
$adobePDFRegPath = "HKCU:\Software\Adobe\Adobe PDF"

if (-not (Test-Path $adobePDFRegPath)) {
    New-Item -Path $adobePDFRegPath -Force | Out-Null
}

# Save original settings so we can restore them when done
$origOutputFolder = (Get-ItemProperty $adobePDFRegPath -Name "OutputFolder" -ErrorAction SilentlyContinue).OutputFolder
$origPrompt       = (Get-ItemProperty $adobePDFRegPath -Name "PromptForPDFFilename" -ErrorAction SilentlyContinue).PromptForPDFFilename

# Point output to our folder and suppress the Save As dialog
Set-ItemProperty $adobePDFRegPath -Name "OutputFolder"           -Value $outputFolder -Type String
Set-ItemProperty $adobePDFRegPath -Name "PromptForPDFFilename"   -Value 0             -Type DWord

# --- Counters and results list for summary ---
$totalProcessed = 0
$totalFlattened  = 0
$results         = @()

try {
    # --- Process each PDF ---
    $acrobatApp = New-Object -ComObject AcroExch.App
    $acrobatApp.Hide()

    foreach ($pdf in $pdfFiles) {
        $filePath = $pdf.FullName
        $savePath = Join-Path $outputFolder $pdf.Name
        Write-Host "Processing: $($pdf.Name)"

        # Remove a previous output file so we can detect fresh creation
        if (Test-Path $savePath) {
            Remove-Item $savePath -Force
        }

        $avDoc = New-Object -ComObject AcroExch.AVDoc

        if (-not $avDoc.Open($filePath, "")) {
            Write-Host "  WARNING: Could not open '$($pdf.Name)'. Skipping."
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null
            continue
        }

        $pdDoc     = $avDoc.GetPDDoc()
        $pageCount = $pdDoc.GetNumPages()

        # Print every page through "Adobe PDF" — this re-renders the
        # document into a new PDF, flattening all interactive content.
        $printOk = $false
        try {
            $printOk = $avDoc.PrintPagesEx(
                0, ($pageCount - 1), 2, $true, $true, $false, $false,
                0,               # bPrintAsImage (0 = vector, preserves quality)
                $printerName, "", ""
            )
        } catch {
            Write-Host "  ERROR: PrintPagesEx failed - $_"
        }

        $status = "Error"

        if ($printOk) {
            # Wait for Distiller to finish writing the output file
            $maxWait = 120
            $elapsed = 0
            while (-not (Test-Path $savePath) -and $elapsed -lt $maxWait) {
                Start-Sleep -Seconds 1
                $elapsed++
            }

            if (Test-Path $savePath) {
                # Let the file finish writing (size stabilises)
                Start-Sleep -Seconds 1
                $s1 = (Get-Item $savePath).Length
                Start-Sleep -Seconds 1
                $s2 = (Get-Item $savePath).Length

                if ($s1 -eq $s2 -and $s1 -gt 0) {
                    Write-Host "  Flattened and saved ($pageCount page(s))."
                    $totalFlattened++
                    $status = "Flattened"
                } else {
                    Write-Host "  WARNING: Output file may be incomplete for '$($pdf.Name)'."
                    $status = "Incomplete"
                }
            } else {
                Write-Host "  ERROR: Output file was not created within the timeout for '$($pdf.Name)'."
            }
        } else {
            Write-Host "  ERROR: Print failed for '$($pdf.Name)'."
        }

        $avDoc.Close($true) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc) | Out-Null

        $totalProcessed++
        $results += [PSCustomObject]@{ FileName = $pdf.Name; Pages = $pageCount; Status = $status }
    }

    # --- Cleanup Acrobat ---
    $acrobatApp.Exit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($acrobatApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
finally {
    # --- Restore original "Adobe PDF" registry settings ---
    if ($null -ne $origOutputFolder) {
        Set-ItemProperty $adobePDFRegPath -Name "OutputFolder" -Value $origOutputFolder -Type String
    } else {
        Remove-ItemProperty $adobePDFRegPath -Name "OutputFolder" -ErrorAction SilentlyContinue
    }
    if ($null -ne $origPrompt) {
        Set-ItemProperty $adobePDFRegPath -Name "PromptForPDFFilename" -Value $origPrompt -Type DWord
    } else {
        Remove-ItemProperty $adobePDFRegPath -Name "PromptForPDFFilename" -ErrorAction SilentlyContinue
    }
}

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
