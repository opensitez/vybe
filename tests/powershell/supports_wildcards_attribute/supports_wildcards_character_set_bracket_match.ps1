# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_character_set_bracket_match
# Wildcard character set brackets [abc] match any one character enclosed within the set
function MatchNodeRevision {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $nodes = @("node-a", "node-b", "node-c", "node-x", "node-z")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matches = @($nodes | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$result = MatchNodeRevision -Pattern "node-[ac]"

if ($result -ne "node-a, node-c") {
    Write-Host "FAIL: character set bracket matching failed: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
