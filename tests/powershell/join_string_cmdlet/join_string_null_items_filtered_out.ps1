# vybe-test: powershell/join_string_cmdlet/join_string_null_items_filtered_out
$list = @('prefix', $null, 'suffix')

# In Join-String, $null items in the pipeline stream are skipped without generating phantom separators
$res = $list | Join-String -Separator ','

if ($res -ne "prefix,suffix") {
    Write-Host "FAIL: expected 'prefix,suffix', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
