# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_string_array_patterns
# SupportsWildcards applied to a string array parameter handles multiple patterns concurrently
function MatchMultiPatterns {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string[]]$Patterns
    )
    process {
        $files = @("test.cs", "test.ps1", "main.go", "config.json")
        $matched = @()
        foreach ($p in $Patterns) {
            $wp = [System.Management.Automation.WildcardPattern]::new($p)
            foreach ($f in $files) {
                if ($wp.IsMatch($f) -and -not $matched.Contains($f)) {
                    $matched += $f
                }
            }
        }
        return $matched -join ", "
    }
}

$results = MatchMultiPatterns -Patterns @("*.cs", "*.ps1")

if ($results -ne "test.cs, test.ps1") {
    Write-Host "FAIL: multiple patterns match failed: '$results'"
    exit 1
}

Write-Host "PASS"
exit 0
