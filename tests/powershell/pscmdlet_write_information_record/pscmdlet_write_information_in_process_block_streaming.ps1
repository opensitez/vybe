# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_in_process_block_streaming
# $PSCmdlet.WriteInformation in the process block emits an InformationRecord per streamed pipeline item
function ProcessInfoEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $PSCmdlet.WriteInformation("processing item: $InputObject", @("ItemTag"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $records = @(10..12 | ProcessInfoEmitter 6>&1)

    if ($records.Count -ne 3) {
        Write-Host "FAIL: expected 3 records, got $($records.Count)"
        exit 1
    }

    for ($i = 0; $i -lt 3; $i++) {
        $expected = "processing item: $(10 + $i)"
        if ($records[$i].MessageData -ne $expected) {
            Write-Host "FAIL: at index ${i}, expected '$expected', got '$($records[$i].MessageData)'"
            exit 1
        }
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
