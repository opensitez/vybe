# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_in_begin_block
# $PSCmdlet.WriteInformation emits InformationRecord instances from within the begin block
function BeginInfoEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin {
        $PSCmdlet.WriteInformation("Begin block initialization", @("Lifecycle"))
    }
    process {
        $InputObject * 2
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $results = @(1..2 | BeginInfoEmitter 6>&1)

    # 1 InformationRecord from begin + 2 output values
    if ($results.Count -ne 3) {
        Write-Host "FAIL: expected 3 stream entries, got $($results.Count)"
        exit 1
    }

    if ($results[0].GetType().Name -ne "InformationRecord") {
        Write-Host "FAIL: first item was not InformationRecord"
        exit 1
    }

    if ($results[0].MessageData -ne "Begin block initialization") {
        Write-Host "FAIL: begin message mismatch"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
