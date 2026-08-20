$files = @("alpha", "bravo", "charlie", "delta", "echo")

$long = $files |
    Where-Object { $_.Length -gt 4 } |
    Sort-Object |
    ForEach-Object { "item: $_" }

$long | ForEach-Object { Write-Output $_ }
