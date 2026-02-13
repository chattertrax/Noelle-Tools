# Noelle-Tools

A collection of utility scripts for everyday document workflows.

## Remove-Pages.ps1

Removes pages containing a specified phrase from all PDF files in a selected folder using Adobe Acrobat Pro.

### Prerequisites

- Windows with PowerShell 5.1 or later
- Adobe Acrobat Pro installed (required for COM automation)

### Setup

1. Open `Remove-Pages.ps1` in a text editor.
2. On line 10, replace `"REPLACE WITH YOUR PHRASE"` with the phrase you want to match against:
   ```powershell
   $Phrase = "your phrase here"
   ```

### Running the Script

#### From PowerShell

```powershell
powershell -File "C:\Path\To\Remove-Pages.ps1"
```

#### Creating a Windows Desktop Shortcut

1. Right-click the desktop and select **New > Shortcut**.
2. In the **location** field, enter:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\Remove-Pages.ps1"
   ```
   Replace `C:\Path\To\Remove-Pages.ps1` with the actual path to the script.
3. Click **Next**, give the shortcut a name (e.g. `Remove PDF Pages`), and click **Finish**.

### Usage

1. Run the script (or double-click the shortcut).
2. A folder selection dialog will appear — browse to the folder containing your PDF files.
3. The script will process every `.pdf` file in that folder:
   - Pages containing the phrase (case-insensitive) are removed.
   - The modified file is saved over the original.
4. Progress and results are printed to the console for each file.

### Behavior Notes

- If no pages in a file match the phrase, the file is left unchanged.
- If **all** pages in a file match, the file is skipped with a warning to avoid creating an empty document.
- Files that cannot be opened are skipped with a warning.
