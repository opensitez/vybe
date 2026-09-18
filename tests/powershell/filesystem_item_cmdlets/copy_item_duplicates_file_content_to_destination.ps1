# vybe-test: powershell/filesystem_item_cmdlets/copy_item_duplicates_file_content_to_destination
# Copy-Item duplicates a source file to a new destination path while retaining content
$src = [System.IO.Path]::GetTempFileName()
$dst = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_cp_dst_$PID.txt"
"test copy payload 123" | Set-Content $src

try {
    Copy-Item -Path $src -Destination $dst

    if (-not (Test-Path $dst -PathType Leaf)) {
        Write-Host "FAIL: destination file does not exist after Copy-Item"
        exit 1
    }

    $dstContent = Get-Content -Path $dst -Raw
    if ($dstContent.Trim() -ne "test copy payload 123") {
        Write-Host "FAIL: unexpected destination content: '$dstContent'"
        exit 1
    }

    # Ensure source still exists
    if (-not (Test-Path $src)) {
        Write-Host "FAIL: source file was removed during Copy-Item"
        exit 1
    }
} finally {
    if (Test-Path $src) { Remove-Item $src -Force }
    if (Test-Path $dst) { Remove-Item $dst -Force }
}

Write-Host "PASS"
exit 0
