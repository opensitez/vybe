# vybe-test: powershell/join_path_cmdlet/join_path_pipeline_streaming
$baseDirs = @("dirA", "dirB")

# Piping base paths into Join-Path joins ChildPath onto each incoming item
$joined = @($baseDirs | Join-Path -ChildPath "manifest.json")

if ($joined.Count -ne 2) {
    Write-Host "FAIL: expected 2 joined paths, got $($joined.Count)"
    exit 1
}

if (-not ($joined[0] -match "^dirA[/|\\]manifest\.json$" -and $joined[1] -match "^dirB[/|\\]manifest\.json$")) {
    Write-Host "FAIL: pipeline joined paths mismatch: @($($joined -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
