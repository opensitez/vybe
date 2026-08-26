# vybe-test: powershell/collections_generic_queue/queue_of_pscustomobjects
$q = [System.Collections.Generic.Queue[pscustomobject]]::new()
$q.Enqueue([pscustomobject]@{ Id = 1; Name = "A" })
$q.Enqueue([pscustomobject]@{ Id = 2; Name = "B" })
$first = $q.Dequeue()
if ($first.Id -ne 1 -or $first.Name -ne "A" -or $q.Count -ne 1) {
    Write-Host "FAIL: Queue of PSCustomObjects failed"
    exit 1
}
Write-Host "PASS"
exit 0
