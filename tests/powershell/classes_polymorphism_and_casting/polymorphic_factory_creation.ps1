# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_factory_creation
class Product { [string]$Category = "General" }
class Book : Product { Book() { $this.Category = "Books" } }
class Electronics : Product { Electronics() { $this.Category = "Electronics" } }
function New-Product([string]$type) {
    if ($type -eq "book") { return [Book]::new() }
    return [Electronics]::new()
}
[Product]$p1 = New-Product "book"
[Product]$p2 = New-Product "elec"
if ($p1.Category -ne "Books" -or $p2.Category -ne "Electronics") {
    Write-Host "FAIL: Polymorphic factory failed"
    exit 1
}
Write-Host "PASS"
exit 0
