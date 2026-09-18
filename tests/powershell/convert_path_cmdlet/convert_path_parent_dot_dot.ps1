# vybe-test: powershell/convert_path_cmdlet/convert_path_parent_dot_dot
# Convert-Path on '..' resolves the parent directory path
$res = Convert-Path ".."

$expectedParent = [System.IO.Directory]::GetParent($PWD.Path).FullName

if ($res -ne $expectedParent) {
    Write-Host "FAIL: parent path conversion failed, expected '$expectedParent', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
