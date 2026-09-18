# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_wildcardpattern_escape_helper
# WildcardPattern.Escape static method escapes wildcard metacharacters with backticks
function EscapeWildcardInput {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        return [System.Management.Automation.WildcardPattern]::Escape($Pattern)
    }
}

$escaped = EscapeWildcardInput -Pattern "report*[2026]?.csv"
$expected = 'report`*`[2026`]`?.csv'

if ($escaped -ne $expected) {
    Write-Host "FAIL: escaped output mismatch: expected '$expected', got '$escaped'"
    exit 1
}

Write-Host "PASS"
exit 0
