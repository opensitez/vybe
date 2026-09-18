# vybe-test: powershell/supports_wildcards_attribute/supports_wildcards_pipeline_binding_streaming
# Parameter decorated with [SupportsWildcards()] accepts streaming patterns from pipeline
function FilterByStreamingPattern {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [SupportsWildcards()]
        [string]$Pattern
    )
    process {
        $targets = @("worker-alpha", "worker-beta", "db-primary")
        $wp = [System.Management.Automation.WildcardPattern]::new($Pattern)
        @($targets | Where-Object { $wp.IsMatch($_) })
    }
}

$streamed = @(@("worker-*", "db-*") | FilterByStreamingPattern)

if ($streamed.Count -ne 3) {
    Write-Host "FAIL: expected 3 streamed wildcard matches, got $($streamed.Count)"
    exit 1
}

$expected = @("worker-alpha", "worker-beta", "db-primary")
for ($i = 0; $i -lt 3; $i++) {
    if ($streamed[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected '$($expected[$i])', got '$($streamed[$i])'"
        exit 1
    }
}

Write-Host "PASS"
exit 0
