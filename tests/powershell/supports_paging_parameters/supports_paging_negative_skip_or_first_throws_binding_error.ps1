# vybe-test: powershell/supports_paging_parameters/supports_paging_negative_skip_or_first_throws_binding_error
# Passing a negative number to -First or -Skip throws ParameterBindingException due to UInt64 type constraint
function TestPagingConstraint {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        "ok"
    }
}

$threwFirst = $false
try {
    TestPagingConstraint -First -5 -ErrorAction Stop
} catch [System.Management.Automation.ParameterBindingException] {
    $threwFirst = $true
}

$threwSkip = $false
try {
    TestPagingConstraint -Skip -1 -ErrorAction Stop
} catch [System.Management.Automation.ParameterBindingException] {
    $threwSkip = $true
}

if (-not $threwFirst -or (-not $threwSkip)) {
    Write-Host "FAIL: negative values for UInt64 paging parameters did not throw ParameterBindingException"
    exit 1
}

Write-Host "PASS"
exit 0
