# vybe-test: powershell/out_variables/out_variable_custom_function
function Get-SampleData {
    [CmdletBinding()]
    param()
    Write-Output "Line1"
    Write-Output "Line2"
}
Get-SampleData -OutVariable lines | Out-Null
if ($lines.Count -ne 2 -or $lines[0] -ne "Line1") {
    Write-Host "FAIL: custom function OutVariable expected Line1, Line2"
    exit 1
}
Write-Host "PASS"
exit 0
