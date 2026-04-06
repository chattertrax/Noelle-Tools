# ============================================================
# Get-FormFields.ps1
# Scans all PDFs in a user-selected input folder and extracts
# the names of every text input field found. Produces a CSV
# report in the chosen output folder listing each field by
# file, field name, field type, and page number.
#
# Uses iText7 (.NET assemblies) — no Acrobat dependency.
# ============================================================

# --- Load iText7 assemblies ---
. "$PSScriptRoot\Includes\Load-iText.ps1"

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
$outputFolder = Select-Folder -Title "Select the OUTPUT folder for the CSV report"
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
$totalFields    = 0
$results        = @()

# --- Process each PDF ---
foreach ($pdf in $pdfFiles) {
    $filePath = $pdf.FullName
    Write-Host "Processing: $($pdf.Name)"

    $pdfReader   = $null
    $pdfDocument = $null

    try {
        $pdfReader   = New-Object iText.Kernel.Pdf.PdfReader($filePath)
        $pdfDocument = New-Object iText.Kernel.Pdf.PdfDocument($pdfReader)
        $form        = [iText.Forms.PdfAcroForm]::GetAcroForm($pdfDocument, $false)

        if (-not $form) {
            Write-Host "  No form fields found."
            $totalProcessed++
            continue
        }

        $fields     = $form.GetAllFormFields()
        $fieldCount = 0

        foreach ($kvp in $fields) {
            $fieldName = $kvp.Key
            $field     = $kvp.Value

            # Filter to text input fields only
            if ($field -isnot [iText.Forms.Fields.PdfTextFormField]) {
                continue
            }

            # Determine which page the field widget is on
            $pageNumber = 0
            $widgets    = $field.GetWidgets()
            if ($widgets -and $widgets.Count -gt 0) {
                $widget = $widgets[0]
                $page   = $widget.GetPage()
                if ($page) {
                    $pageNumber = $pdfDocument.GetPageNumber($page)
                }
            }

            $fieldCount++
            $results += [PSCustomObject]@{
                FileName   = $pdf.Name
                FieldName  = $fieldName
                FieldType  = "Text"
                PageNumber = $pageNumber
            }
        }

        $totalFields += $fieldCount
        Write-Host "  Found $fieldCount text field(s)."

    } catch {
        Write-Host "  ERROR: $($_.Exception.Message)"
    } finally {
        if ($pdfDocument) { $pdfDocument.Close() }
        elseif ($pdfReader) { $pdfReader.Close() }
    }

    $totalProcessed++
}

Write-Host ""
Write-Host "Done. All files processed."

# --- Export CSV to the output folder with a timestamp filename ---
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath   = Join-Path $outputFolder "$timestamp.csv"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# --- Summary popup with option to view the CSV ---
$answer = [System.Windows.Forms.MessageBox]::Show(
    "Total Processed: $totalProcessed`nTotal Text Fields Found: $totalFields`n`nWould you like to view the summary?",
    "Get Form Fields - Complete",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
    Invoke-Item $csvPath
}
