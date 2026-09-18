# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_single_pattern_asterisk_match
# Parameter decorated with [SupportsWildcards()] matching arbitrary characters using asterisk wildcard
function SearchLogEntries {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $pool = @("app_error.log", "app_audit.log", "sys_metric.csv", "app_trace.log")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern, [System.Management.Automation.WildcardOptions]::IgnoreCase)
        $matches = @($pool | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$result = SearchLogEntries -Pattern "app_*.log"

if ($result -ne "app_error.log, app_audit.log, app_trace.log") {
    Write-Host "FAIL: asterisk pattern match mismatch: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
