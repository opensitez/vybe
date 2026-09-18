# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_attribute_instantiation
# SupportsWildcardsAttribute can be instantiated and its type verified
$attr = [System.Management.Automation.SupportsWildcardsAttribute]::new()

if ($attr -eq $null) {
    Write-Host "FAIL: SupportsWildcardsAttribute was null"
    exit 1
}

$expectedType = "System.Management.Automation.SupportsWildcardsAttribute"
if ($attr.GetType().FullName -ne $expectedType) {
    Write-Host "FAIL: expected type '$expectedType', got '$($attr.GetType().FullName)'"
    exit 1
}

Write-Host "PASS"
exit 0
