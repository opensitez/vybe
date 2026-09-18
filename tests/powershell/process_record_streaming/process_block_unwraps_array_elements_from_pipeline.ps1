# vybe-test: powershell/process_record_streaming/process_block_unwraps_array_elements_from_pipeline
# The pipeline operator unwraps an array into individual scalar items passed to process one-by-one
$script:receivedTypes = @()

function AuditReceivedTypes {
    process {
        $script:receivedTypes += $_.GetType().Name
    }
}

$inputArray = @(10, 20, 30)
$inputArray | AuditReceivedTypes

# Must receive 3 separate Int32 values, not 1 Object[] array
if ($script:receivedTypes.Count -ne 3) {
    Write-Host "FAIL: expected 3 scalar calls, got $($script:receivedTypes.Count)"
    exit 1
}

foreach ($typeName in $script:receivedTypes) {
    if ($typeName -ne "Int32") {
        Write-Host "FAIL: expected Int32 scalar type, got $typeName"
        exit 1
    }
}

Write-Host "PASS"
exit 0
