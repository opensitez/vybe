# vybe-test: powershell/whitespace_and_line_rules/empty_line_in_pipeline
$script = "1,2,3 |`n`n Measure-Object -Sum | Select-Object -ExpandProperty Sum"
$errors = @()
$values = [System.Management.Automation.PSParser]::Tokenize($script, [ref]$errors)
if ($errors.Count -ne 0) {
    Write-Host 'FAIL: blank line in pipeline introduced parser errors'
    exit 1
}

if ((Invoke-Expression $script) -ne 6) {
    Write-Host 'FAIL: empty-line pipeline output incorrect'
    exit 1
}

Write-Host 'PASS'
exit 0
