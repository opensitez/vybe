# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_in_clean_block
# $PSCmdlet.WriteInformation emits InformationRecord entries from within the PowerShell 7.3+ clean block
function CleanInfoEmitter {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    process {
        $InputObject
    }
    clean {
        $PSCmdlet.WriteInformation("clean block finalization", @("Teardown"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $results = @(1 | CleanInfoEmitter 6>&1)

    # 1 data + 1 InformationRecord from clean
    if ($results.Count -ne 2) {
        Write-Host "FAIL: expected 2 entries, got $($results.Count)"
        exit 1
    }

    $infoItem = $results | Where-Object { $_ -is [System.Management.Automation.InformationRecord] }
    if ($infoItem -eq $null -or $infoItem.MessageData -ne "clean block finalization") {
        Write-Host "FAIL: clean block InformationRecord mismatch"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
