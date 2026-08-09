# vybe-test: powershell/ordered_hashtables/ordered_hashtable_iteration
$h = [ordered]@{ X = 10; Y = 20; Z = 30 }
$collected = @()
foreach ($entry in $h.GetEnumerator()) {
    $collected += "$($entry.Key)=$($entry.Value)"
}
$str = $collected -join ";"
if ($str -ne "X=10;Y=20;Z=30") {
    Write-Host "FAIL: enumerator order expected X=10;Y=20;Z=30, got $str"
    exit 1
}
Write-Host "PASS"
exit 0
