# vybe-test: powershell/engine_strict_mode/strict_mode_v2_hashtable_index_lookup_missing_key_returns_null
Set-StrictMode -Version 2.0

$hash = @{ Title = "Guide" }

# Key distinction: strict mode v2.0 prohibits DOT property access on missing keys,
# but INDEX lookup $hash["Key"] for a non-existent key evaluates to $null without throwing.
$caught = $false
$val = $null
try {
    $val = $hash["MissingKeyName"]
} catch {
    $caught = $true
}

if ($caught) {
    Write-Host "FAIL: hashtable index lookup for missing key threw under v2.0"
    exit 1
}

if ($null -ne $val) {
    Write-Host "FAIL: expected missing key index lookup to be `$null, got: $val"
    exit 1
}

Write-Host "PASS"
exit 0
