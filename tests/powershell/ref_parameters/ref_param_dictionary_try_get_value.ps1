# vybe-test: powershell/ref_parameters/ref_param_dictionary_try_get_value
$dict = [System.Collections.Generic.Dictionary[string, string]]::new()
$dict["Version"] = "1.0"
$v = ""
$found = $dict.TryGetValue("Version", [ref]$v)
if (-not $found -or $v -ne "1.0") {
    Write-Host "FAIL: Dictionary TryGetValue via [ref] expected true and '1.0'"
    exit 1
}
Write-Host "PASS"
exit 0
