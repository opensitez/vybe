# vybe-test: powershell/classes_static_constructors/static_constructor_with_environment_variable_read
class EnvReader {
    static [string]$PathSep
    static EnvReader() {
        [EnvReader]::PathSep = [System.IO.Path]::DirectorySeparatorChar.ToString()
    }
}
if ([EnvReader]::PathSep.Length -ne 1) {
    Write-Host "FAIL: Static directory separator check failed"
    exit 1
}
Write-Host "PASS"
exit 0
