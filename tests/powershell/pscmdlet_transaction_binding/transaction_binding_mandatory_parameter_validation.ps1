# vybe-test: powershell/pscmdlet_write_error_record/transaction_binding_mandatory_parameter_validation
# Mandatory parameter declarations function properly when declared in a SupportsTransactions function
function CommitDatabaseTxn {
    [CmdletBinding(SupportsTransactions = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ConnectionUri
    )
    process {
        return "Connected:$ConnectionUri"
    }
}

# 1. Successful execution when mandatory parameter is provided
$res = CommitDatabaseTxn -ConnectionUri "tcp://db.local:5432"

if ($res -ne "Connected:tcp://db.local:5432") {
    Write-Host "FAIL: mandatory parameter execution failed: '$res'"
    exit 1
}

# 2. Reflection confirms Mandatory = true on parameter attribute
$cmd = Get-Command CommitDatabaseTxn
$paramInfo = $cmd.Parameters["ConnectionUri"]
$isMandatory = $paramInfo.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] } | Select-Object -ExpandProperty Mandatory

if (-not $isMandatory) {
    Write-Host "FAIL: ConnectionUri was not marked Mandatory"
    exit 1
}

Write-Host "PASS"
exit 0
