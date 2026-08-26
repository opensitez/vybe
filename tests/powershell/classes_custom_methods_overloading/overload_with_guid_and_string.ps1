# vybe-test: powershell/classes_custom_methods_overloading/overload_with_guid_and_string
class KeyLookup {
    [string]Find([string]$str) { return "FoundString:$str" }
    [string]Find([guid]$g) { return "FoundGuid:$($g.ToString())" }
}
$kl = [KeyLookup]::new()
$g = [guid]::Parse("11111111-1111-1111-1111-111111111111")
$r1 = $kl.Find("my-key")
$r2 = $kl.Find($g)
if ($r1 -ne "FoundString:my-key" -or $r2 -ne "FoundGuid:11111111-1111-1111-1111-111111111111") {
    Write-Host "FAIL: Guid vs String overload failed"
    exit 1
}
Write-Host "PASS"
exit 0
