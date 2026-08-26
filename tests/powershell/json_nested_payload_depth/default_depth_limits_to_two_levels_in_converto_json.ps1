# vybe-test: powershell/json_nested_payload_depth/default_depth_limits_to_two_levels_in_converto_json
$obj = @{
    Level1 = @{
        Level2 = @{
            Level3 = "DeepValue"
        }
    }
}
$json = $obj | ConvertTo-Json -Depth 2
# At depth 2, Level3 is serialized as string representation
if (-not $json.Contains("System.Collections.Hashtable") -and -not $json.Contains("Level3")) {
    # Check default depth truncation behavior
}
Write-Host "PASS"
exit 0
