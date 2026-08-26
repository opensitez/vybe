# vybe-test: powershell/pipeline_custom_enumerator_streaming/custom_enumerator_emitting_custom_objects
class PersonRecord { [string]$Name; PersonRecord([string]$n) { $this.Name = $n } }
class PeopleGen : System.Collections.IEnumerable {
    [System.Collections.IEnumerator]GetEnumerator() { return [PeopleEnum]::new() }
}
class PeopleEnum : System.Collections.IEnumerator {
    [string[]]$Names = @("Alice", "Bob")
    [int]$Idx = -1
    [object] get_Current() { return [PersonRecord]::new($this.Names[$this.Idx]) }
    [bool] MoveNext() { $this.Idx++; return ($this.Idx -lt $this.Names.Length) }
    [void] Reset() { $this.Idx = -1 }
}
$pg = [PeopleGen]::new()
$names = @($pg | ForEach-Object { $_.Name })
if ($names[0] -ne "Alice" -or $names[1] -ne "Bob") {
    Write-Host "FAIL: Custom enumerator emitting custom objects failed"
    exit 1
}
Write-Host "PASS"
exit 0
