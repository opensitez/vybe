# vybe-test: powershell/format_list_engine/list_format_enumeration_limit_ellipsis
$data = [pscustomobject]@{ Elements = 1..10 }

# $FormatEnumerationLimit bounds the number of collection elements rendered in list view
$prev = $FormatEnumerationLimit
$FormatEnumerationLimit = 2

$output = $data | Format-List -Property Elements | Out-String

$FormatEnumerationLimit = $prev

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

$hasEllipsis = ($output -match "\u2026") -or ($output -match "\.\.\.")
if (-not $hasEllipsis) {
    Write-Host "FAIL: list view did not truncate collection with ellipsis: $output"
    exit 1
}

Write-Host "PASS"
exit 0
