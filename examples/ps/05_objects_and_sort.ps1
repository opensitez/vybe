$users = @(
    [pscustomobject]@{Name="Ana"; Age=34},
    [pscustomobject]@{Name="Ben"; Age=28},
    [pscustomobject]@{Name="Maya"; Age=41}
)

$users |
    Sort-Object -Property Age |
    ForEach-Object {
        Write-Output "$($_.Name) is $($_.Age) years old"
    }
