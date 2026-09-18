# vybe-test: powershell/process_record_streaming/process_block_with_pipeline_parameter_by_property_name
# Declaring [Parameter(ValueFromPipelineByPropertyName=$true)] binds matching object properties to parameters in process
function FormatUserInfo {
    param(
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Username,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [int]$UserId
    )
    process {
        "${UserId}:${Username}"
    }
}

$inputObjects = @(
    [PSCustomObject]@{ Username = "admin"; UserId = 101 },
    [PSCustomObject]@{ Username = "operator"; UserId = 102 }
)

$results = @($inputObjects | FormatUserInfo)

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 results, got $($results.Count)"
    exit 1
}

if ($results[0] -ne "101:admin" -or $results[1] -ne "102:operator") {
    Write-Host "FAIL: property binding mismatch: $($results -join ', ')"
    exit 1
}

Write-Host "PASS"
exit 0
