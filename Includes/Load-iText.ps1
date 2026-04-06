# Load-iText.ps1
# Dot-source this script to load all iText7 assemblies into the current session.
# Usage: . "$PSScriptRoot\..\Includes\Load-iText.ps1"

$includesDir = $PSScriptRoot

# Load all DLLs from the Includes folder
$dlls = Get-ChildItem -Path $includesDir -Filter '*.dll' -ErrorAction Stop
foreach ($dll in $dlls) {
    try {
        [System.Reflection.Assembly]::LoadFrom($dll.FullName) | Out-Null
    } catch {
        Write-Warning "Failed to load $($dll.Name): $($_.Exception.Message)"
    }
}
