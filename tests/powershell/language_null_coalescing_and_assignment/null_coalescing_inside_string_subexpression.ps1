# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_inside_string_subexpression
$name = $null
$msg = "Hello $( $name ?? 'Guest' )!"
if ($msg -ne "Hello Guest!") {
    Write-Host "FAIL: ?? in string subexpression failed, got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
