# vybe-test: powershell/filesystem_item_cmdlets/copy_item_force_overwrites_existing_destination_file
# Copy-Item with -Force overwrites an already existing destination file with the source file content
$src = [System.IO.Path]::GetTempFileName()
$dst = [System.IO.Path]::GetTempFileName()

"newer source data" | Set-Content $src
"initial destination data" | Set-Content $dst

try {
    Copy-Item -Path $src -Destination $dst -Force

    $content = Get-Content -Path $dst -Raw
    if ($content.Trim() -ne "newer source data") {
        Write-Host "FAIL: destination file was not overwritten with source data"
        exit 1
    }
} finally {
    if (Test-Path $src) { Remove-Item $src -Force }
    if (Test-Path $dst) { Remove-Item $dst -Force }
}

Write-Host "PASS"
exit 0
