# vybe-test: powershell/control_flow/switch_file_input_stream
# Switch statement with -File parameter processes input line by line from an external file
$tmp = [System.IO.Path]::GetTempFileName()

try {
    Set-Content -Path $tmp -Value @("header_line", "target_line", "footer_line")

    $matched = @()
    switch -File $tmp {
        "target_line" { $matched += "found:$_" }
        default       { $matched += "other:$_" }
    }

    if ($matched.Count -ne 3) {
        Write-Host "FAIL: expected 3 processed lines, got $($matched.Count)"
        exit 1
    }

    if ($matched[1] -ne "found:target_line") {
        Write-Host "FAIL: expected 'found:target_line', got '$($matched[1])'"
        exit 1
    }
} finally {
    if (Test-Path $tmp) {
        Remove-Item $tmp -Force
    }
}

Write-Host "PASS"
exit 0
