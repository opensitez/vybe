# vybe-test: powershell/collections_linked_list/linked_list_contains_check
$ll = [System.Collections.Generic.LinkedList[string]]::new([string[]]@("apple", "banana"))
if (-not $ll.Contains("apple") -or $ll.Contains("orange")) { Write-Host "FAIL: Contains check failed"; exit 1 }
Write-Host "PASS"; exit 0
