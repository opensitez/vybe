# vybe-test: powershell/supports_paging_parameters/supports_paging_new_total_count_helper_method
# PagingParameters.NewTotalCount(ulong, double) creates a total count record object with an Accuracy note property
function TestNewTotalCountHelper {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        if ($PSCmdlet.PagingParameters.IncludeTotalCount) {
            $PSCmdlet.PagingParameters.NewTotalCount(1500, 0.98)
        }
    }
}

$record = TestNewTotalCountHelper -IncludeTotalCount

if ($record -ne 1500) {
    Write-Host "FAIL: total count value mismatch, expected 1500, got: $record"
    exit 1
}

$accuracy = $record.PSObject.Properties["Accuracy"].Value
if ($accuracy -ne 0.98) {
    Write-Host "FAIL: accuracy property mismatch: '$accuracy'"
    exit 1
}

Write-Host "PASS"
exit 0
