# vybe-test: powershell/classes_static_constructors/static_constructor_with_string_builder_construction
class Banner {
    static [string]$Header
    static Banner() {
        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append("===").Append("SYSTEM").Append("===")
        [Banner]::Header = $sb.ToString()
    }
}
if ([Banner]::Header -ne "===SYSTEM===") {
    Write-Host "FAIL: Static StringBuilder construction failed"
    exit 1
}
Write-Host "PASS"
exit 0
