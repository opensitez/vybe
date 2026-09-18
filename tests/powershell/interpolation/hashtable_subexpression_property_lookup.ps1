# vybe-test: powershell/interpolation/hashtable_subexpression_property_lookup
# Hashtable index lookups inside a subexpression $() correctly interpolate into a double-quoted string
$config = @{
    Environment = "production"
    Port = 8080
    Active = $true
}

$status = "Running on $($config['Environment']) host at port $($config['Port'])"

$expected = "Running on production host at port 8080"
if ($status -ne $expected) {
    Write-Host "FAIL: expected '$expected', got '$status'"
    exit 1
}

Write-Host "PASS"
exit 0
