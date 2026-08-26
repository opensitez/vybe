# vybe-test: powershell/collections_generic_queue/peek_without_removing
$q = [System.Collections.Generic.Queue[string]]::new()
$q.Enqueue("stay")
$peeked = $q.Peek()
if ($peeked -ne "stay" -or $q.Count -ne 1) {
    Write-Host "FAIL: Queue Peek should not remove element"
    exit 1
}
Write-Host "PASS"
exit 0
