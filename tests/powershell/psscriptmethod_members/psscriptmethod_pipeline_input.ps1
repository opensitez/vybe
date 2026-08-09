# vybe-test: powershell/psscriptmethod_members/psscriptmethod_pipeline_input
$items = @([pscustomobject]@{ N = 2 }, [pscustomobject]@{ N = 3 })
$items | ForEach-Object { $_ | Add-Member -MemberType ScriptMethod -Name "Cube" -Value { $this.N * $this.N * $this.N } }
if ($items[0].Cube() -ne 8 -or $items[1].Cube() -ne 27) {
    Write-Host "FAIL: pipeline attached PSScriptMethod expected Cube() 8, 27"
    exit 1
}
Write-Host "PASS"
exit 0
