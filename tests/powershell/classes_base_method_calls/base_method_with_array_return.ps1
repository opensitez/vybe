# vybe-test: powershell/classes_base_method_calls/base_method_with_array_return
class BaseProvider {
    [string[]]GetDefaultTags() { return @("v1", "stable") }
}
class CustomProvider : BaseProvider {
    [string[]]GetAllTags() {
        $baseTags = ([BaseProvider]$this).GetDefaultTags()
        return @($baseTags) + "custom"
    }
}
$cp = [CustomProvider]::new()
$tags = $cp.GetAllTags()
if ($tags.Length -ne 3 -or $tags[0] -ne "v1" -or $tags[2] -ne "custom") {
    Write-Host "FAIL: Base method array return failed"
    exit 1
}
Write-Host "PASS"
exit 0
