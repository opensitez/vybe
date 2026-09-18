# vybe-test: powershell/filesystem_item_cmdlets/rename_item_changes_filename_preserving_parent_directory
# Rename-Item changes the leaf name of a file while leaving it inside the same directory
$dir = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_rn_dir_$PID"
New-Item -Path $dir -ItemType Directory -Force | Out-Null
$orig = Join-Path $dir "before_rename.txt"
"content before rename" | Set-Content $orig

try {
    Rename-Item -Path $orig -NewName "after_rename.txt"

    if (Test-Path $orig) {
        Write-Host "FAIL: original file still exists after Rename-Item"
        exit 1
    }

    $renamed = Join-Path $dir "after_rename.txt"
    if (-not (Test-Path $renamed -PathType Leaf)) {
        Write-Host "FAIL: renamed file does not exist at expected path"
        exit 1
    }

    $content = Get-Content -Path $renamed -Raw
    if ($content.Trim() -ne "content before rename") {
        Write-Host "FAIL: renamed file content altered"
        exit 1
    }
} finally {
    if (Test-Path $dir) {
        Remove-Item $dir -Recurse -Force
    }
}

Write-Host "PASS"
exit 0
