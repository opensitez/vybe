# vybe-test: powershell/supports_paging_parameters/supports_paging_psboundparameters_contains_explicit_keys
# $PSBoundParameters contains entries only for the paging parameters explicitly provided by the caller
function CheckBoundPagingKeys {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        $hasFirst = $PSBoundParameters.ContainsKey("First")
        $hasSkip = $PSBoundParameters.ContainsKey("Skip")
        $hasTotal = $PSBoundParameters.ContainsKey("IncludeTotalCount")
        return "First:$hasFirst,Skip:$hasSkip,Total:$hasTotal"
    }
}

$callOnlyFirst = CheckBoundPagingKeys -First 10
$callSkipAndTotal = CheckBoundPagingKeys -Skip 20 -IncludeTotalCount

if ($callOnlyFirst -ne "First:True,Skip:False,Total:False") {
    Write-Host "FAIL: call with only First mismatch: '$callOnlyFirst'"
    exit 1
}

if ($callSkipAndTotal -ne "First:False,Skip:True,Total:True") {
    Write-Host "FAIL: call with Skip and Total mismatch: '$callSkipAndTotal'"
    exit 1
}

Write-Host "PASS"
exit 0
