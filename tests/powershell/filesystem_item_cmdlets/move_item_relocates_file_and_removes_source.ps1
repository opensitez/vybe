# vybe-test: powershell/filesystem_item_cmdlets/move_item_relocates_file_and_removes_source
# Move-Item moves a file to a new target path, leaving no file at the original source location
$src = [System.IO.Path]::GetTempFileName()
$dst = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_mv_dst_$PID.txt"
"move payload alpha" | Set-Content $src

try {
    Move-Item -Path $src -Destination $dst

    if (Test-Path $src) {
        Write-Host "FAIL: source file still exists after Move-Item"
        exit 1
    }

    if (-not (Test-Path $dst -PathType Leaf)) {
        Write-Host "FAIL: destination file does not exist after Move-Item"
        exit 1
    }

    $content = Get-Content -Path $dst -Raw
    if ($content.Trim() -ne "move payload alpha") {
        Write-Host "FAIL: destination file content mismatch"
        exit 1
    }
} finally {
    if (Test-Path $src) { Remove-Item $src -Force }
    if (Test-Path $dst) { Remove-Item $dst -Force }
}

Write-Host "PASS"
exit 0
