# vybe-test: powershell/filesystem_item_cmdlets/remove_item_recurse_deletes_directory_tree
# Remove-Item with -Recurse removes a directory containing subdirectories and files
$baseDir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_rm_tree_$PID"
$subDir = Join-Path $baseDir "sub"
New-Item -Path $subDir -ItemType Directory -Force | Out-Null
"leaf file" | Set-Content (Join-Path $subDir "child.txt")

if (-not (Test-Path (Join-Path $subDir "child.txt"))) {
    Write-Host "FAIL: precondition failed: child file not created"
    exit 1
}

Remove-Item -Path $baseDir -Recurse -Force

if (Test-Path $baseDir) {
    Write-Host "FAIL: base directory still exists after Remove-Item -Recurse"
    exit 1
}

Write-Host "PASS"
exit 0
