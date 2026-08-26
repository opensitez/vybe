# vybe-test: powershell/classes_custom_methods_overloading/overload_with_hashtable_and_dictionary
class DictHandler {
    [int]CountKeys([hashtable]$ht) { return $ht.Count }
    [int]CountKeys([System.Collections.Generic.Dictionary[string, int]]$dict) { return $dict.Count }
}
$dh = [DictHandler]::new()
$ht = @{ a = 1; b = 2 }
$d = [System.Collections.Generic.Dictionary[string, int]]::new()
$d.Add("x", 10); $d.Add("y", 20); $d.Add("z", 30)
if ($dh.CountKeys($ht) -ne 2 -or $dh.CountKeys($d) -ne 3) {
    Write-Host "FAIL: Hashtable vs Dictionary overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
