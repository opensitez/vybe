# vybe-test: powershell/select_string_cmdlet/select_string_quiet_file_returns_boolean_false
# When querying a file with -Quiet and no matches exist, Select-String returns boolean $false
$found = Select-String -Path "Cargo.toml" -Pattern "completely_fabricated_token_xyz_9988" -Quiet

if ($found -ne $false) {
    Write-Host "FAIL: Select-String -Quiet on non-matching file expected `$false, got: $found"
    exit 1
}

Write-Host "PASS"
exit 0
