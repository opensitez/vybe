# vybe-test: powershell/intrinsic_collection_overloads/where_on_hashtable_enumerator
$hash = @{
    Apple  = 100
    Banana = 250
    Cherry = 500
}

# In PowerShell, invoking .Where() on a hashtable enumerator filters DictionaryEntry objects
$res = $hash.GetEnumerator().Where({ $_.Value -ge 200 }, 'Default')

if ($null -eq $res) {
    Write-Host "FAIL: .Where() on hashtable enumerator returned `$null"
    exit 1
}

if ($res.Count -ne 2) {
    Write-Host "FAIL: expected 2 entries with Value >= 200, got $($res.Count)"
    exit 1
}

# Verify entries have expected keys and values
$keys = @($res[0].Key, $res[1].Key)
if (-not ($keys -contains "Banana" -and $keys -contains "Cherry")) {
    Write-Host "FAIL: expected Banana and Cherry, got @($($keys -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
