# vybe-test: powershell/hashtables/hashtable_clone_shallow_copy
# The Clone() method creates a separate copy of the hashtable so modifications do not affect original
$original = @{
    Primary = "Alpha"
    Secondary = "Beta"
}

$cloned = $original.Clone()

# Modify the clone
$cloned["Tertiary"] = "Gamma"
$cloned["Primary"] = "ModifiedAlpha"

if ($original.Count -ne 2) {
    Write-Host "FAIL: original hashtable count changed after mutating clone"
    exit 1
}

if ($original["Primary"] -ne "Alpha") {
    Write-Host "FAIL: original key was overwritten by clone modification"
    exit 1
}

if ($original.ContainsKey("Tertiary")) {
    Write-Host "FAIL: key added to clone appeared in original"
    exit 1
}

Write-Host "PASS"
exit 0
