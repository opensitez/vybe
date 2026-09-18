# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_single_pattern_question_mark_match
# Wildcard pattern matches exactly one single character per question mark symbol
function MatchPortIdentifier {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $ports = @("port-1", "port-12", "port-80", "port-443", "port-9")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matches = @($ports | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$result = MatchPortIdentifier -Pattern "port-??"

if ($result -ne "port-12, port-80") {
    Write-Host "FAIL: question mark wildcard matching failed: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
