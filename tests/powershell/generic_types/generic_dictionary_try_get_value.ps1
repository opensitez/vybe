# vybe-test: powershell/generic_types/generic_dictionary_try_get_value
$dict = [System.Collections.Generic.Dictionary[string, int]]::new()
$dict["K"] = 777
$val = 0
$found = $dict.TryGetValue("K", [ref]$val)
if (-not $found -or $val -ne 777) {
    Write-Host "FAIL: TryGetValue expected true and 777, got found=$found val=$val"
    exit 1
}
Write-Host "PASS"
exit 0
