# vybe-test: powershell/process_record_streaming/process_block_comma_operator_preserves_array_as_single_item
# Prefixing an array with the unary comma operator prevents pipeline unwrapping, sending the array as 1 object
$processCalls = 0
$receivedTypeName = ""

function CheckSingleArrayObject {
    process {
        $script:processCalls++
        $script:receivedTypeName = $_.GetType().Name
    }
}

$nestedArray = @(1, 2, 3, 4, 5)
, $nestedArray | CheckSingleArrayObject

if ($processCalls -ne 1) {
    Write-Host "FAIL: expected exactly 1 process invocation for wrapped array, got $processCalls"
    exit 1
}

if ($receivedTypeName -ne "Object[]") {
    Write-Host "FAIL: expected Object[] type, got '$receivedTypeName'"
    exit 1
}

Write-Host "PASS"
exit 0
