# vybe-test: powershell/collections_generic_stack/trypeek_success_and_empty
$s = [System.Collections.Generic.Stack[string]]::new()
$val = ""
$bad = $s.TryPeek([ref]$val)
$s.Push("data")
$ok = $s.TryPeek([ref]$val)
if ($bad -or -not $ok -or $val -ne "data") {
    Write-Host "FAIL: TryPeek failed"
    exit 1
}
Write-Host "PASS"
exit 0
