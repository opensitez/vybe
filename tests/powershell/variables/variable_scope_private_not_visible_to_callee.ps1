# vybe-test: powershell/variables/variable_scope_private_not_visible_to_callee
# A variable defined in the private: scope is not visible to child scopes or called functions
$private:secretVar = "superSecret"

function CheckChildVisibility {
    return $secretVar
}

$visibleInCallee = CheckChildVisibility

if ($visibleInCallee -ne $null) {
    Write-Host "FAIL: private variable leaked into child function scope: '$visibleInCallee'"
    exit 1
}

# But it should remain visible in the defining scope
if ($private:secretVar -ne "superSecret") {
    Write-Host "FAIL: private variable not accessible in defining scope"
    exit 1
}

Write-Host "PASS"
exit 0
