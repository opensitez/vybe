# vybe-test: powershell/supports_paging_parameters/supports_paging_pscmdlet_paging_parameters_object
# $PSCmdlet.PagingParameters provides a non-null instance of System.Management.Automation.PagingParameters
function TestPagingParametersObject {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.GetType().FullName
    }
}

$typeName = TestPagingParametersObject

if ($typeName -ne "System.Management.Automation.PagingParameters") {
    Write-Host "FAIL: unexpected PagingParameters type: '$typeName'"
    exit 1
}

Write-Host "PASS"
exit 0
