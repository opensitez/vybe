# vybe-test: powershell/generic_types/generic_dictionary_instantiation
$dict = [System.Collections.Generic.Dictionary[string, int]]::new()
$dict.Add("One", 1)
if ($dict["One"] -ne 1) {
    Write-Host "FAIL: Dictionary[string,int] key 'One' expected 1, got $($dict['One'])"
    exit 1
}
Write-Host "PASS"
exit 0
