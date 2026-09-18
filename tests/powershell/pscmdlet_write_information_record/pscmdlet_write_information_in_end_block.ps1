# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_in_end_block
# $PSCmdlet.WriteInformation emits summary InformationRecord entries from within the end block
function EndInfoEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin { $total = 0 }
    process { $total += $InputObject }
    end {
        $PSCmdlet.WriteInformation("aggregate result: $total", @("Summary"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $records = @(5..8 | EndInfoEmitter 6>&1)

    # 5 + 6 + 7 + 8 = 26
    if ($records.Count -ne 1) {
        Write-Host "FAIL: expected 1 end block record, got $($records.Count)"
        exit 1
    }

    if ($records[0].MessageData -ne "aggregate result: 26") {
        Write-Host "FAIL: end message mismatch: '$($records[0].MessageData)'"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
