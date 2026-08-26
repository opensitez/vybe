# vybe-test: powershell/collections_tuples_and_valuetuples/tuple_in_hashtable_key
$t = [System.Tuple]::Create("x", "y")
$ht = @{ $t = "origin" }
$lookup = [System.Tuple]::Create("x", "y")
if ($ht[$lookup] -ne "origin") {
    Write-Host "FAIL: Tuple as hashtable key lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
