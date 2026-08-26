# vybe-test: powershell/parameters_validate_length/validatelength_whitespace_included_in_length_count
function Set-Spaced {
    param([ValidateLength(5, 10)][string]$Str)
    return $Str.Length
}
$res = Set-Spaced -Str "  a  " # length 5
if ($res -ne 5) {
    Write-Host "FAIL: Whitespace counting in ValidateLength failed"
    exit 1
}
Write-Host "PASS"
exit 0
