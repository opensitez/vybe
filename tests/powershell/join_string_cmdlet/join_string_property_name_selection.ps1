# vybe-test: powershell/join_string_cmdlet/join_string_property_name_selection
$servers = @(
    [pscustomobject]@{ Hostname = "web01"; IP = "10.0.0.1" },
    [pscustomobject]@{ Hostname = "web02"; IP = "10.0.0.2" }
)

# -Property extracts the property value from pipeline objects before joining
$res = $servers | Join-String -Property Hostname -Separator ','

if ($res -ne "web01,web02") {
    Write-Host "FAIL: expected 'web01,web02', got '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
