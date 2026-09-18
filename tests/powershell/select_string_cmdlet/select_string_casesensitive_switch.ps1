# vybe-test: powershell/select_string_cmdlet/select_string_casesensitive_switch
# -CaseSensitive enforces exact casing requirements
$mismatch = "UPPERCASE_TEXT" | Select-String -Pattern "uppercase_text" -CaseSensitive
$matching = "UPPERCASE_TEXT" | Select-String -Pattern "UPPERCASE_TEXT" -CaseSensitive

if ($null -ne $mismatch) {
    Write-Host "FAIL: -CaseSensitive unexpectedly matched different casing"
    exit 1
}

if ($null -eq $matching) {
    Write-Host "FAIL: -CaseSensitive failed to match identical casing"
    exit 1
}

Write-Host "PASS"
exit 0
