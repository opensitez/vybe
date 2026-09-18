# vybe-test: powershell/intrinsic_collection_overloads/foreach_type_cast_conversion
$stringNumbers = @("10", "20", "30")

# .ForEach([type]) performs bulk fast casting of each element into the target type
$converted = $stringNumbers.ForEach([int])

if ($null -eq $converted -or $converted.Count -ne 3) {
    Write-Host "FAIL: .ForEach([int]) expected 3 items, got $($converted.Count)"
    exit 1
}

# Verify type of elements is System.Int32
foreach ($item in $converted) {
    if (-not ($item -is [int])) {
        Write-Host "FAIL: element $item was not cast to [int], type is $($item.GetType().FullName)"
        exit 1
    }
}

# Concrete value verification: numeric addition vs string concatenation
$sum = $converted[0] + $converted[1] + $converted[2]
if ($sum -ne 60) {
    Write-Host "FAIL: numeric addition expected 60, got $sum"
    exit 1
}

Write-Host "PASS"
exit 0
