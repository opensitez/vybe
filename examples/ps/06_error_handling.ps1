function Get-NumberFromText {
    param([string]$Text)
    return [int]$Text
}

try {
    $value = Get-NumberFromText -Text "99"
    Write-Output "Parsed value: $value"

    $value = Get-NumberFromText -Text "n/a"
    Write-Output "This line will not run"
}
catch {
    Write-Output "Parse failed: $($_.Exception.Message)"
}
finally {
    Write-Output "Cleanup complete."
}
