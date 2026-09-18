# vybe-test: powershell/new_object_cmdlet/new_object_generic_dictionary_string_int
# New-Object with a quoted generic type name creates a Dictionary[string,int]
$dict = New-Object "System.Collections.Generic.Dictionary[string,int]"
$dict["alpha"] = 1
$dict["beta"] = 2

if ($dict["alpha"] -ne 1 -or $dict["beta"] -ne 2) {
    Write-Host "FAIL: dictionary value mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
