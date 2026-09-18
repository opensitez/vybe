# vybe-test: powershell/test_path_cmdlet/test_path_pathtype_container_directory_true
# Test-Path with -PathType Container returns $true for existing directories
$res = Test-Path "tests" -PathType Container

if ($res -ne $true) {
    Write-Host "FAIL: expected `$true for directory with -PathType Container, got: $res"
    exit 1
}

Write-Host "PASS"
exit 0
