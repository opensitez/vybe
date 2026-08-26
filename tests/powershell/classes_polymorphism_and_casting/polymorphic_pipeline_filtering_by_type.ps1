# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_pipeline_filtering_by_type
class ItemBase {}
class FileItem : ItemBase { [string]$Name; FileItem([string]$n) { $this.Name = $n } }
class DirItem : ItemBase { [string]$Path; DirItem([string]$p) { $this.Path = $p } }
[ItemBase[]]$items = @([FileItem]::new("f1"), [DirItem]::new("/d1"), [FileItem]::new("f2"))
$files = @($items | Where-Object { $_ -is [FileItem] })
if ($files.Count -ne 2 -or $files[0].Name -ne "f1" -or $files[1].Name -ne "f2") {
    Write-Host "FAIL: Polymorphic pipeline type filtering failed"
    exit 1
}
Write-Host "PASS"
exit 0
