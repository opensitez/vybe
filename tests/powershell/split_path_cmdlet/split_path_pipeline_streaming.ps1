# vybe-test: powershell/split_path_cmdlet/split_path_pipeline_streaming
$paths = @("/var/log/syslog", "/etc/hosts")

# Piping paths into Split-Path streams each processed result
$leaves = @($paths | Split-Path -Leaf)

if ($leaves.Count -ne 2) {
    Write-Host "FAIL: expected 2 leaves, got $($leaves.Count)"
    exit 1
}

if ($leaves[0] -ne "syslog" -or $leaves[1] -ne "hosts") {
    Write-Host "FAIL: expected @('syslog', 'hosts'), got: @($($leaves -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
