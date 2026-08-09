# vybe-test: powershell/out_variables/out_variable_nested_function
function Inner-Func {
    [CmdletBinding()]
    param()
    Write-Output "InnerOutput"
}
function Outer-Func {
    Inner-Func -OutVariable script:outFromInner | Out-Null
}
Outer-Func
if ($script:outFromInner[0] -ne "InnerOutput") {
    Write-Host "FAIL: nested function OutVariable expected InnerOutput"
    exit 1
}
Write-Host "PASS"
exit 0
