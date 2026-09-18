# vybe-test: powershell/pscmdlet_write_information_record/pscmdlet_write_information_preserves_hashtable_message_data
# Passing a hashtable to $PSCmdlet.WriteInformation stores the hashtable directly in MessageData with key lookup
function EmitHashtableInfo {
    [CmdletBinding()]
    param()
    process {
        $dict = @{
            SessionId = 4096
            ClusterState = "Converged"
        }
        $PSCmdlet.WriteInformation($dict, @("Cluster"))
    }
}

$oldPreference = $InformationPreference
$InformationPreference = "Continue"

try {
    $record = (EmitHashtableInfo 6>&1)[0]
    $payload = $record.MessageData

    if ($payload["SessionId"] -ne 4096 -or $payload["ClusterState"] -ne "Converged") {
        Write-Host "FAIL: hashtable MessageData keys mismatch"
        exit 1
    }
} finally {
    $InformationPreference = $oldPreference
}

Write-Host "PASS"
exit 0
