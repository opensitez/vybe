# vybe-test: powershell/exceptions_try_catch_multiple_typed_blocks/catch_block_preserves_error_record_category
$cat = $null
try {
    [int]::Parse("xyz")
} catch [System.FormatException] {
    $cat = $_.CategoryInfo.Category
}
if ($cat -eq $null) {
    Write-Host "FAIL: CategoryInfo preserved in typed catch check failed"
    exit 1
}
Write-Host "PASS"
exit 0
