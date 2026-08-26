# vybe-test: powershell/collections_generic_queue/trypeek_success_and_empty
$q = [System.Collections.Generic.Queue[string]]::new()
$val = ""
$bad = $q.TryPeek([ref]$val)
$q.Enqueue("data")
$ok = $q.TryPeek([ref]$val)
if ($bad -or -not $ok -or $val -ne "data") {
    Write-Host "FAIL: TryPeek failed"
    exit 1
}
Write-Host "PASS"
exit 0
