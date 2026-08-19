# vybe-test: powershell/boolean_value_rules/basic
$enabled = $true -and (-not $false)
$disabled = $false -or $null

if (-not $enabled) {
    Write-Host 'FAIL: true-and-not-false should be true'
    exit 1
}

if ($disabled) {
    Write-Host 'FAIL: false-or-null should be false'
    exit 1
}

Write-Host 'PASS'
exit 0
