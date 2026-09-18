# vybe-test: powershell/filesystem_item_cmdlets/move_item_force_overwrites_target_file
# Move-Item with -Force overwrites an existing file at the destination path
$src = [System.IO.Path]::GetTempFileName()
$dst = [System.IO.Path]::GetTempFileName()
"source overwrite content" | Set-Content $src
"existing destination content" | Set-Content $dst

try {
    Move-Item -Path $src -Destination $dst -Force

    if (Test-Path $src) {
        Write-Host "FAIL: source file still exists after Move-Item -Force"
        exit 1
    }

    $content = Get-Content -Path $dst -Raw
    if ($content.Trim() -ne "source overwrite content") {
        Write-Host "FAIL: destination file was not overwritten with source content"
        exit 1
    }
} finally {
    if (Test-Path $src) { Remove-Item $src -Force }
    if (Test-Path $dst) { Remove-Item $dst -Force }
}

Write-Host "PASS"
exit 0
