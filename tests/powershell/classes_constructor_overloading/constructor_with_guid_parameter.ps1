# vybe-test: powershell/classes_constructor_overloading/constructor_with_guid_parameter
class Session {
    [guid]$Id
    Session() { $this.Id = [guid]::NewGuid() }
    Session([guid]$id) { $this.Id = $id }
}
$g = [guid]::Parse("11111111-2222-3333-4444-555555555555")
$s = [Session]::new($g)
if ($s.Id -ne $g) {
    Write-Host "FAIL: Guid constructor failed"
    exit 1
}
Write-Host "PASS"
exit 0
