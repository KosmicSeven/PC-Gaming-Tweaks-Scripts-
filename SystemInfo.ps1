<#
.SYNOPSIS
    Exports system diagnostic information to the Desktop.

.DESCRIPTION
    Runs DXDiag, Get-ComputerInfo, and systeminfo and saves the output
    to text files on the current user's Desktop.

.OUTPUTS
    PC_DXDiag.txt       - DirectX diagnostic report (Desktop)
    PC_ComputerInfo.txt - PowerShell ComputerInfo report (Desktop)
    PC_SystemInfo.txt   - Windows systeminfo report (Desktop)

.NOTES
    Run as Administrator for full output.
#>

$Desktop = [Environment]::GetFolderPath("Desktop")

Write-Host "Running DXDiag..." -ForegroundColor Cyan
dxdiag /t "$Desktop\PC_DXDiag.txt"

Write-Host "Running Get-ComputerInfo..." -ForegroundColor Cyan
Get-ComputerInfo | Out-File "$Desktop\PC_ComputerInfo.txt"

Write-Host "Running systeminfo..." -ForegroundColor Cyan
systeminfo > "$Desktop\PC_SystemInfo.txt"

Write-Host ""
Write-Host "Done. Files saved to: $Desktop" -ForegroundColor Green
