# vybe-test: powershell/variables/variable_allscope_inherited_to_new_scopes
# A variable defined with the AllScope option is duplicated and visible in child scopes
$allScopeName = "allScopeTest_$PID"

try {
    New-Variable -Name $allScopeName -Value "shared_state" -Option AllScope -Force

    function ReadAllScopeVar {
        return (Get-Variable -Name $allScopeName).Value
    }

    $childVal = ReadAllScopeVar
    if ($childVal -ne "shared_state") {
        Write-Host "FAIL: AllScope variable not seen in child function, got: '$childVal'"
        exit 1
    }
} finally {
    Remove-Variable -Name $allScopeName -Force -ErrorAction SilentlyContinue
}

Write-Host "PASS"
exit 0
