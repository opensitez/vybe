# vybe-test: powershell/collections_generic_queue/copyto_target_array
$q = [System.Collections.Generic.Queue[string]]::new()
$q.Enqueue("x"); $q.Enqueue("y")
[string[]]$target = New-Object string[] 2
$q.CopyTo($target, 0)
if ($target[0] -ne "x" -or $target[1] -ne "y") {
    Write-Host "FAIL: CopyTo on Queue failed"
    exit 1
}
Write-Host "PASS"
exit 0
