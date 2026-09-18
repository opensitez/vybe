# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_wildcardpattern_contains_wildcard_detection
# WildcardPattern.ContainsWildcardCharacters static method identifies strings with wildcards
function CheckContainsWildcard {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        return [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($Pattern)
    }
}

$hasWildcard = CheckContainsWildcard -Pattern "server-*.corp"
$isLiteral = CheckContainsWildcard -Pattern "server-prod-01.corp"

if (-not $hasWildcard) {
    Write-Host "FAIL: failed to detect wildcard character in 'server-*.corp'"
    exit 1
}

if ($isLiteral) {
    Write-Host "FAIL: falsely detected wildcard in literal string 'server-prod-01.corp'"
    exit 1
}

Write-Host "PASS"
exit 0
