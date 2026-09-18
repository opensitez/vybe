# vybe-test: powershell/filesystem_item_cmdlets/new_item_creates_intermediate_directories_with_force
# New-Item creates missing intermediate parent directories when creating a deeply nested file with -Force
$base = Join-Path ([System.IO.Path]::GetTempPath()) "vybe_deep_ni_$PID"
$nested = Join-Path $base "level1/level2/nested_file.txt"

try {
    New-Item -Path $nested -ItemType File -Value "deep payload" -Force | Out-Null

    if (-not (Test-Path $nested -PathType Leaf)) {
        Write-Host "FAIL: nested file was not created"
        exit 1
    }

    $content = Get-Content -Path $nested -Raw
    if ($content.Trim() -ne "deep payload") {
        Write-Host "FAIL: unexpected content in nested file: $content"
        exit 1
    }
} finally {
    if (Test-Path $base) {
        Remove-Item $base -Recurse -Force
    }
}

Write-Host "PASS"
exit 0
