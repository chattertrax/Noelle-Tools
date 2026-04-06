# ============================================================
# Start-PDFTools.ps1
# WPF GUI that consolidates three PDF operations using iText7:
#   1. Remove Pages  — delete pages containing a phrase
#   2. Flatten Pages  — convert form fields to static content
#   3. Remove Fields  — selectively remove text fields by name
#
# Configuration is persisted in Memory.json alongside this script.
# ============================================================

# --- Load iText7 assemblies ---
. "$PSScriptRoot\Includes\Load-iText.ps1"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# Config helpers
# ============================================================
$script:ConfigPath = Join-Path $PSScriptRoot "Memory.json"

function Load-Config {
    if (Test-Path $script:ConfigPath) {
        return Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
    }
    return [PSCustomObject]@{
        FieldRanges  = @()
        RemovePhrase = "Information Missing on the Document"
    }
}

function Save-Config {
    param ($Config)
    $Config | ConvertTo-Json -Depth 3 | Set-Content $script:ConfigPath -Encoding UTF8
}

# ============================================================
# Field-range expansion: returns a hashtable for O(1) lookup
# ============================================================
function Expand-FieldRanges {
    param ($FieldRanges)
    $names = @{}
    foreach ($r in $FieldRanges) {
        for ($i = [int]$r.Start; $i -le [int]$r.End; $i++) {
            $names["$($r.Prefix)$i"] = $true
        }
    }
    return $names
}

# ============================================================
# Tool 1 — Remove Pages (iText7)
# ============================================================
function Invoke-RemovePages {
    param ($InputFolder, $OutputFolder, $Phrase, $GenerateCsv)

    $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf" -File
    if ($pdfFiles.Count -eq 0) { Write-Host "No PDF files found."; return }

    Write-Host "Remove Pages: Found $($pdfFiles.Count) file(s). Phrase: `"$Phrase`""
    Write-Host ""

    $results        = @()
    $totalProcessed = 0
    $totalModified  = 0

    foreach ($pdf in $pdfFiles) {
        Write-Host "Processing: $($pdf.Name)"
        $outPath = Join-Path $OutputFolder $pdf.Name

        $pdfDoc = $null
        try {
            # First pass: identify pages containing the phrase
            $reader = New-Object iText.Kernel.Pdf.PdfReader($pdf.FullName)
            $pdfDoc = New-Object iText.Kernel.Pdf.PdfDocument($reader)
            $pageCount    = $pdfDoc.GetNumberOfPages()
            $pagesToRemove = @()

            for ($p = 1; $p -le $pageCount; $p++) {
                $page = $pdfDoc.GetPage($p)
                $text = [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page)
                if ($text -match [regex]::Escape($Phrase)) {
                    $pagesToRemove += $p
                }
            }
            $pdfDoc.Close(); $pdfDoc = $null

            if ($pagesToRemove.Count -eq 0) {
                Write-Host "  No matching pages. Copying unchanged."
                Copy-Item -LiteralPath $pdf.FullName -Destination $outPath -Force
                $results += [PSCustomObject]@{ FileName = $pdf.Name; PagesRemoved = 0; Status = "Unchanged" }
            }
            elseif ($pagesToRemove.Count -eq $pageCount) {
                Write-Host "  WARNING: All $pageCount page(s) match. Skipping to avoid empty document."
                $results += [PSCustomObject]@{ FileName = $pdf.Name; PagesRemoved = 0; Status = "Skipped (all match)" }
            }
            else {
                # Second pass: open with writer and remove pages
                $reader2 = New-Object iText.Kernel.Pdf.PdfReader($pdf.FullName)
                $writer2 = New-Object iText.Kernel.Pdf.PdfWriter($outPath)
                $pdfDoc  = New-Object iText.Kernel.Pdf.PdfDocument($reader2, $writer2)

                foreach ($p in ($pagesToRemove | Sort-Object -Descending)) {
                    $pdfDoc.RemovePage($p)
                }
                $pdfDoc.Close(); $pdfDoc = $null

                $remaining = $pageCount - $pagesToRemove.Count
                Write-Host "  Removed $($pagesToRemove.Count) page(s). $remaining remaining."
                $totalModified++
                $results += [PSCustomObject]@{ FileName = $pdf.Name; PagesRemoved = $pagesToRemove.Count; Status = "Modified" }
            }
        }
        catch {
            Write-Host "  ERROR: $($_.Exception.Message)"
            $results += [PSCustomObject]@{ FileName = $pdf.Name; PagesRemoved = 0; Status = "Error" }
        }
        finally {
            if ($pdfDoc) { $pdfDoc.Close() }
        }
        $totalProcessed++
    }

    Write-Host ""
    Write-Host "Done. Processed: $totalProcessed, Modified: $totalModified"
    Show-Summary -Results $results -OutputFolder $OutputFolder -ToolName "RemovePages" -GenerateCsv $GenerateCsv `
                 -SummaryText "Total Processed: $totalProcessed`nTotal Modified: $totalModified"
}

# ============================================================
# Tool 2 — Flatten Pages (iText7)
# ============================================================
function Invoke-FlattenPages {
    param ($InputFolder, $OutputFolder, $GenerateCsv)

    $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf" -File
    if ($pdfFiles.Count -eq 0) { Write-Host "No PDF files found."; return }

    Write-Host "Flatten Pages: Found $($pdfFiles.Count) file(s)."
    Write-Host ""

    $results        = @()
    $totalProcessed = 0
    $totalFlattened = 0

    foreach ($pdf in $pdfFiles) {
        Write-Host "Processing: $($pdf.Name)"
        $outPath = Join-Path $OutputFolder $pdf.Name

        $pdfDoc = $null
        try {
            $reader = New-Object iText.Kernel.Pdf.PdfReader($pdf.FullName)
            $writer = New-Object iText.Kernel.Pdf.PdfWriter($outPath)
            $pdfDoc = New-Object iText.Kernel.Pdf.PdfDocument($reader, $writer)

            $form = [iText.Forms.PdfAcroForm]::GetAcroForm($pdfDoc, $false)
            if ($form) {
                $form.FlattenFields()
                Write-Host "  Flattened ($($pdfDoc.GetNumberOfPages()) page(s))."
                $totalFlattened++
                $results += [PSCustomObject]@{ FileName = $pdf.Name; Pages = $pdfDoc.GetNumberOfPages(); Status = "Flattened" }
            }
            else {
                Write-Host "  No form fields. Copied unchanged."
                $results += [PSCustomObject]@{ FileName = $pdf.Name; Pages = $pdfDoc.GetNumberOfPages(); Status = "No Fields" }
            }
            $pdfDoc.Close(); $pdfDoc = $null
        }
        catch {
            Write-Host "  ERROR: $($_.Exception.Message)"
            $results += [PSCustomObject]@{ FileName = $pdf.Name; Pages = 0; Status = "Error" }
        }
        finally {
            if ($pdfDoc) { $pdfDoc.Close() }
        }
        $totalProcessed++
    }

    Write-Host ""
    Write-Host "Done. Processed: $totalProcessed, Flattened: $totalFlattened"
    Show-Summary -Results $results -OutputFolder $OutputFolder -ToolName "FlattenPages" -GenerateCsv $GenerateCsv `
                 -SummaryText "Total Processed: $totalProcessed`nTotal Flattened: $totalFlattened"
}

# ============================================================
# Tool 3 — Remove Fields (iText7)
# ============================================================
function Invoke-RemoveFields {
    param ($InputFolder, $OutputFolder, $FieldRanges, $GenerateCsv)

    $expandedNames = Expand-FieldRanges -FieldRanges $FieldRanges
    if ($expandedNames.Count -eq 0) {
        Write-Host "No field ranges configured. Nothing to remove."
        return
    }

    $pdfFiles = Get-ChildItem -Path $InputFolder -Filter "*.pdf" -File
    if ($pdfFiles.Count -eq 0) { Write-Host "No PDF files found."; return }

    Write-Host "Remove Fields: Found $($pdfFiles.Count) file(s). Targeting $($expandedNames.Count) field name(s)."
    Write-Host ""

    $results        = @()
    $totalProcessed = 0
    $totalModified  = 0

    foreach ($pdf in $pdfFiles) {
        Write-Host "Processing: $($pdf.Name)"
        $outPath = Join-Path $OutputFolder $pdf.Name

        $pdfDoc = $null
        try {
            $reader = New-Object iText.Kernel.Pdf.PdfReader($pdf.FullName)
            $writer = New-Object iText.Kernel.Pdf.PdfWriter($outPath)
            $pdfDoc = New-Object iText.Kernel.Pdf.PdfDocument($reader, $writer)

            $form = [iText.Forms.PdfAcroForm]::GetAcroForm($pdfDoc, $false)
            $removedCount = 0

            if ($form) {
                $fields = $form.GetAllFormFields()

                # Collect names to remove (cannot modify collection during iteration)
                $toRemove = @()
                foreach ($kvp in $fields) {
                    if ($kvp.Value -is [iText.Forms.Fields.PdfTextFormField] -and $expandedNames.ContainsKey($kvp.Key)) {
                        $toRemove += $kvp.Key
                    }
                }

                foreach ($name in $toRemove) {
                    $form.RemoveField($name)
                    $removedCount++
                }
            }

            $pdfDoc.Close(); $pdfDoc = $null

            if ($removedCount -gt 0) {
                Write-Host "  Removed $removedCount field(s)."
                $totalModified++
            }
            else {
                Write-Host "  No matching fields found."
            }
            $results += [PSCustomObject]@{ FileName = $pdf.Name; FieldsRemoved = $removedCount; Status = if ($removedCount -gt 0) { "Modified" } else { "Unchanged" } }
        }
        catch {
            Write-Host "  ERROR: $($_.Exception.Message)"
            $results += [PSCustomObject]@{ FileName = $pdf.Name; FieldsRemoved = 0; Status = "Error" }
        }
        finally {
            if ($pdfDoc) { $pdfDoc.Close() }
        }
        $totalProcessed++
    }

    Write-Host ""
    Write-Host "Done. Processed: $totalProcessed, Modified: $totalModified"
    Show-Summary -Results $results -OutputFolder $OutputFolder -ToolName "RemoveFields" -GenerateCsv $GenerateCsv `
                 -SummaryText "Total Processed: $totalProcessed`nTotal Modified: $totalModified"
}

# ============================================================
# CSV export and summary
# ============================================================
function Show-Summary {
    param ($Results, $OutputFolder, $ToolName, $GenerateCsv, $SummaryText)

    $csvPath = $null
    if ($GenerateCsv -and $Results.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $csvPath   = Join-Path $OutputFolder "${ToolName}_${timestamp}.csv"
        $Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    }

    $msg = "$SummaryText"
    if ($csvPath) {
        $msg += "`n`nWould you like to view the summary?"
        $answer = [System.Windows.MessageBox]::Show($msg, "PDF Tools - Complete", "YesNo", "Information")
        if ($answer -eq "Yes") { Invoke-Item $csvPath }
    }
    else {
        [System.Windows.MessageBox]::Show($msg, "PDF Tools - Complete", "OK", "Information") | Out-Null
    }
}

# ============================================================
# WPF Window
# ============================================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PDF Tools" Width="720" Height="640"
        ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Folders -->
        <GroupBox Grid.Row="0" Header="Folders" Padding="8" Margin="0,0,0,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="95"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="80"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Text="Input Folder:" VerticalAlignment="Center"/>
                <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtInput" Margin="4,0" VerticalContentAlignment="Center" Height="26"/>
                <Button    Grid.Row="0" Grid.Column="2" x:Name="btnInput" Content="Browse..." Height="26"/>

                <TextBlock Grid.Row="2" Grid.Column="0" Text="Output Folder:" VerticalAlignment="Center"/>
                <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtOutput" Margin="4,0" VerticalContentAlignment="Center" Height="26"/>
                <Button    Grid.Row="2" Grid.Column="2" x:Name="btnOutput" Content="Browse..." Height="26"/>
            </Grid>
        </GroupBox>

        <!-- Configuration -->
        <GroupBox Grid.Row="1" Header="Configuration" Padding="8" Margin="0,0,0,8">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Text="Field ranges to remove:" Margin="0,0,0,4"/>

                <DataGrid Grid.Row="1" x:Name="dgRanges"
                          AutoGenerateColumns="False" CanUserAddRows="False"
                          HeadersVisibility="Column" SelectionMode="Single"
                          VerticalScrollBarVisibility="Auto" MinHeight="120">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Prefix" Binding="{Binding Prefix}" Width="*"/>
                        <DataGridTextColumn Header="Start"  Binding="{Binding Start}"  Width="80"/>
                        <DataGridTextColumn Header="End"    Binding="{Binding End}"    Width="80"/>
                    </DataGrid.Columns>
                </DataGrid>

                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,6,0,0">
                    <Button x:Name="btnAddRow"    Content="Add Row"    Width="80" Margin="0,0,6,0"/>
                    <Button x:Name="btnRemoveRow" Content="Remove Row" Width="90" Margin="0,0,6,0"/>
                    <Button x:Name="btnSave"      Content="Save Config" Width="90"/>
                </StackPanel>

                <DockPanel Grid.Row="4">
                    <TextBlock Text="Remove Phrase:" DockPanel.Dock="Left" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBox x:Name="txtPhrase" VerticalContentAlignment="Center" Height="26"/>
                </DockPanel>
            </Grid>
        </GroupBox>

        <!-- Tool selection and Run -->
        <GroupBox Grid.Row="2" Header="Run Tool" Padding="8">
            <StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                    <RadioButton x:Name="rbRemovePages"  Content="Remove Pages"  IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
                    <RadioButton x:Name="rbFlatten"       Content="Flatten Pages" Margin="0,0,18,0" VerticalAlignment="Center"/>
                    <RadioButton x:Name="rbRemoveFields"  Content="Remove Fields" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                    <CheckBox x:Name="chkCsv" Content="Generate CSV report" IsChecked="True" VerticalAlignment="Center"/>
                </StackPanel>
                <Button x:Name="btnRun" Content="Run" Height="34" FontWeight="Bold"
                        Background="#0078D7" Foreground="White"/>
            </StackPanel>
        </GroupBox>
    </Grid>
</Window>
"@

# --- Parse XAML and find controls ---
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtInput       = $window.FindName("txtInput")
$txtOutput      = $window.FindName("txtOutput")
$btnInput       = $window.FindName("btnInput")
$btnOutput      = $window.FindName("btnOutput")
$dgRanges       = $window.FindName("dgRanges")
$btnAddRow      = $window.FindName("btnAddRow")
$btnRemoveRow   = $window.FindName("btnRemoveRow")
$btnSave        = $window.FindName("btnSave")
$txtPhrase      = $window.FindName("txtPhrase")
$rbRemovePages  = $window.FindName("rbRemovePages")
$rbFlatten      = $window.FindName("rbFlatten")
$rbRemoveFields = $window.FindName("rbRemoveFields")
$chkCsv         = $window.FindName("chkCsv")
$btnRun         = $window.FindName("btnRun")

# --- DataGrid backing collection ---
$script:RangeRows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$dgRanges.ItemsSource = $script:RangeRows

# --- Load config into UI ---
$config = Load-Config
foreach ($r in $config.FieldRanges) {
    $script:RangeRows.Add([PSCustomObject]@{ Prefix = $r.Prefix; Start = [int]$r.Start; End = [int]$r.End })
}
$txtPhrase.Text = $config.RemovePhrase

# ============================================================
# Event handlers
# ============================================================
function Browse-Folder {
    param ($Title)
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description  = $Title
    $dlg.UseDescriptionForTitle = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.SelectedPath
    }
    return $null
}

$btnInput.Add_Click({
    $path = Browse-Folder -Title "Select INPUT folder containing PDF files"
    if ($path) { $txtInput.Text = $path }
})

$btnOutput.Add_Click({
    $path = Browse-Folder -Title "Select OUTPUT folder for processed PDFs"
    if ($path) { $txtOutput.Text = $path }
})

$btnAddRow.Add_Click({
    $script:RangeRows.Add([PSCustomObject]@{ Prefix = "field"; Start = 0; End = 0 })
})

$btnRemoveRow.Add_Click({
    $sel = $dgRanges.SelectedItem
    if ($sel) { $script:RangeRows.Remove($sel) | Out-Null }
})

$btnSave.Add_Click({
    $ranges = @()
    foreach ($row in $script:RangeRows) {
        $ranges += [PSCustomObject]@{ Prefix = $row.Prefix; Start = [int]$row.Start; End = [int]$row.End }
    }
    $cfg = [PSCustomObject]@{
        FieldRanges  = $ranges
        RemovePhrase = $txtPhrase.Text
    }
    Save-Config -Config $cfg
    [System.Windows.MessageBox]::Show("Configuration saved.", "PDF Tools", "OK", "Information") | Out-Null
})

$btnRun.Add_Click({
    # --- Validate ---
    $inputFolder  = $txtInput.Text.Trim()
    $outputFolder = $txtOutput.Text.Trim()

    if (-not $inputFolder -or -not (Test-Path $inputFolder)) {
        [System.Windows.MessageBox]::Show("Please select a valid input folder.", "PDF Tools - Error", "OK", "Error") | Out-Null
        return
    }
    if (-not $outputFolder) {
        [System.Windows.MessageBox]::Show("Please select an output folder.", "PDF Tools - Error", "OK", "Error") | Out-Null
        return
    }
    if ($inputFolder.TrimEnd('\') -eq $outputFolder.TrimEnd('\')) {
        [System.Windows.MessageBox]::Show("Input and output folders must be different.", "PDF Tools - Error", "OK", "Error") | Out-Null
        return
    }
    if (-not (Test-Path $outputFolder)) {
        New-Item -Path $outputFolder -ItemType Directory | Out-Null
    }

    $generateCsv = $chkCsv.IsChecked

    # --- Save config before running ---
    $ranges = @()
    foreach ($row in $script:RangeRows) {
        $ranges += [PSCustomObject]@{ Prefix = $row.Prefix; Start = [int]$row.Start; End = [int]$row.End }
    }
    $cfg = [PSCustomObject]@{
        FieldRanges  = $ranges
        RemovePhrase = $txtPhrase.Text
    }
    Save-Config -Config $cfg

    # --- Dispatch ---
    if ($rbRemovePages.IsChecked) {
        Invoke-RemovePages -InputFolder $inputFolder -OutputFolder $outputFolder -Phrase $txtPhrase.Text -GenerateCsv $generateCsv
    }
    elseif ($rbFlatten.IsChecked) {
        Invoke-FlattenPages -InputFolder $inputFolder -OutputFolder $outputFolder -GenerateCsv $generateCsv
    }
    elseif ($rbRemoveFields.IsChecked) {
        Invoke-RemoveFields -InputFolder $inputFolder -OutputFolder $outputFolder -FieldRanges $cfg.FieldRanges -GenerateCsv $generateCsv
    }
})

# --- Show window ---
$window.ShowDialog() | Out-Null
