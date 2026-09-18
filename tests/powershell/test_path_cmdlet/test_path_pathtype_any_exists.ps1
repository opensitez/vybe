# vybe-test: powershell/test_path_cmdlet/test_path_pathtype_any_exists
# -PathType Any returns $true if the path exists as either a file or a container
$dirExists  = Test-Path "tests" -PathType Any
$fileExists = Test-Path "Cargo.toml" -PathType Any

if ($dirExists -ne $true -or $fileExists -ne $true) {
    Write-Host "FAIL: -PathType Any failed on existing items, got dir=$dirExists, file=$fileExists"
    exit 1
}

Write-Host "PASS"
exit 0
