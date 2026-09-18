# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_character_range_bracket_match
# Wildcard character ranges [0-9] match numeric characters within specified boundary
function MatchBuildNumber {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $builds = @("build-1", "build-5", "build-8", "build-k", "build-omega")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matches = @($builds | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

$result = MatchBuildNumber -Pattern "build-[0-5]"

if ($result -ne "build-1, build-5") {
    Write-Host "FAIL: range bracket matching failed: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
