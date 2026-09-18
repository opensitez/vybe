# vybe-test: powershell/compare_object_cmdlet/compare_object_passthru_returns_original_types
# -PassThru outputs the original unwrapped objects with a dynamic SideIndicator NoteProperty
$diff = Compare-Object @(10, 20) @(20, 30) -PassThru

if ($diff.Count -ne 2) {
    Write-Host "FAIL: expected 2 elements, got $($diff.Count)"
    exit 1
}

if (-not ($diff[0] -is [int])) {
    Write-Host "FAIL: expected original type [int], got: $($diff[0].GetType().FullName)"
    exit 1
}

if ($null -eq $diff[0].SideIndicator) {
    Write-Host "FAIL: dynamic SideIndicator NoteProperty missing on PassThru object"
    exit 1
}

Write-Host "PASS"
exit 0
