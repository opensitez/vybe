# vybe-test: powershell/test_path_cmdlet/test_path_pathtype_container_file_false
# Test-Path with -PathType Container returns $false for a file
$res = Test-Path "Cargo.toml" -PathType Container

if ($res -ne $false) {
    Write-Host "FAIL: expected `$false for file tested with -PathType Container, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
