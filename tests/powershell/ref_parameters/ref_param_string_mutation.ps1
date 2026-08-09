# vybe-test: powershell/ref_parameters/ref_param_string_mutation
function Append-Text([ref]$str, [string]$suffix) {
    $str.Value = $str.Value + $suffix
}
$txt = "Base"
Append-Text ([ref]$txt) "_Suffix"
if ($txt -ne "Base_Suffix") {
    Write-Host "FAIL: [ref] string append expected Base_Suffix, got $txt"
    exit 1
}
Write-Host "PASS"
exit 0
