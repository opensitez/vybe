# vybe-test: powershell/filesystem_item_cmdlets/copy_item_recurse_copies_directory_hierarchy
# Copy-Item with -Recurse copies an entire directory hierarchy including all nested files
$srcDir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_cp_src_$PID"
$subDir = Join-Path $srcDir "subfolder"
$dstDir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_cp_dstdir_$PID"

New-Item -Path $subDir -ItemType Directory -Force | Out-Null
"nested content" | Set-Content (Join-Path $subDir "data.txt")

try {
    Copy-Item -Path $srcDir -Destination $dstDir -Recurse

    $copiedChild = Join-Path $dstDir "subfolder/data.txt"
    if (-not (Test-Path $copiedChild -PathType Leaf)) {
        Write-Host "FAIL: copied nested child file does not exist at '$copiedChild'"
        exit 1
    }

    $content = Get-Content -Path $copiedChild -Raw
    if ($content.Trim() -ne "nested content") {
        Write-Host "FAIL: content mismatch in recursively copied child file"
        exit 1
    }
} finally {
    if (Test-Path $srcDir) { Remove-Item $srcDir -Recurse -Force }
    if (Test-Path $dstDir) { Remove-Item $dstDir -Recurse -Force }
}

Write-Host "PASS"
exit 0
