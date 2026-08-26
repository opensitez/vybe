# vybe-test: powershell/parameters_validate_not_null_or_empty/whitespace_string_is_not_empty
function Check-Spaces {
    param([ValidateNotNullOrEmpty()][string]$Text)
    return $Text.Length
}
$res = Check-Spaces -Text "   " # whitespace string has length 3, not empty
if ($res -ne 3) {
    Write-Host "FAIL: Whitespace string failed ValidateNotNullOrEmpty, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
