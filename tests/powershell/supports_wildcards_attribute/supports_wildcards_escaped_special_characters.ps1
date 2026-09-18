# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_escaped_special_characters
# Escaping wildcard characters using the PowerShell backtick matches literal asterisks and brackets
function MatchLiteralSymbols {
    [CmdletBinding()]
    param(
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $pool = @("file*.txt", "file1.txt", "file2.txt")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        $matches = @($pool | Where-Object { $wp.IsMatch($_) })
        return $matches -join ", "
    }
}

# Literal asterisk escaped via backtick in single quotes
$result = MatchLiteralSymbols -Pattern 'file`*.txt'

if ($result -ne "file*.txt") {
    Write-Host "FAIL: escaped asterisk match failed: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
