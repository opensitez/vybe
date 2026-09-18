# vybe-test: powershell/join_string_cmdlet/join_string_direct_input_object_parameter
# Join-String accepts -InputObject directly without requiring pipeline piping
$res = Join-String -InputObject @('x', 'y', 'z') -Separator '.'

if ($res -ne "x.y.z") {
    Write-Host "FAIL: direct -InputObject failed, expected 'x.y.z', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
