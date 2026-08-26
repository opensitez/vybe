# vybe-test: powershell/type_char_classification_methods/is_control_null_and_newline
$nul = [char]0
$nl = [char]10
$x = [char]'x'
if (-not [char]::IsControl($nul) -or -not [char]::IsControl($nl) -or [char]::IsControl($x)) {
    Write-Host "FAIL: IsControl check failed"
    exit 1
}
Write-Host "PASS"
exit 0
