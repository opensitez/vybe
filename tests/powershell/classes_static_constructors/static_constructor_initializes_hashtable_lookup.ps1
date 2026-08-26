# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_hashtable_lookup
class LookupTable {
    static [hashtable]$Codes
    static LookupTable() {
        [LookupTable]::Codes = @{
            200 = "OK"
            404 = "NotFound"
            500 = "InternalError"
        }
    }
}
$ok = [LookupTable]::Codes[200]
$nf = [LookupTable]::Codes[404]
if ($ok -ne "OK" -or $nf -ne "NotFound") {
    Write-Host "FAIL: Static hashtable lookup failed"
    exit 1
}
Write-Host "PASS"
exit 0
