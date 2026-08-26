# vybe-test: powershell/pipeline_begin_process_end_blocks/pipeline_with_string_builder_in_begin_end
function Concat-Strings {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][string]$Word)
    begin { $sb = [System.Text.StringBuilder]::new() }
    process { $null = $sb.Append($Word).Append(" ") }
    end { return $sb.ToString().TrimEnd() }
}
$res = "The", "quick", "brown", "fox" | Concat-Strings
if ($res -ne "The quick brown fox") {
    Write-Host "FAIL: StringBuilder in begin/end block failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
