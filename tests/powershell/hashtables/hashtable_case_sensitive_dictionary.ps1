# vybe-test: powershell/hashtables/hashtable_case_sensitive_dictionary
# A Hashtable constructed with StringComparer.Ordinal maintains distinct keys for different cases
$caseSensitive = [System.Collections.Hashtable]::new([System.StringComparer]::Ordinal)

$caseSensitive["ITEM"] = "UPPER"
$caseSensitive["item"] = "lower"

if ($caseSensitive.Count -ne 2) {
    Write-Host "FAIL: expected count 2 for distinct cased keys, got $($caseSensitive.Count)"
    exit 1
}

if ($caseSensitive["ITEM"] -ne "UPPER" -or $caseSensitive["item"] -ne "lower") {
    Write-Host "FAIL: case-sensitive key values collided or altered"
    exit 1
}

Write-Host "PASS"
exit 0
