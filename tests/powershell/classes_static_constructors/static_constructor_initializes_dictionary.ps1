# vybe-test: powershell/classes_static_constructors/static_constructor_initializes_dictionary
class DictStatic {
    static [System.Collections.Generic.Dictionary[string, int]]$Weights
    static DictStatic() {
        [DictStatic]::Weights = [System.Collections.Generic.Dictionary[string, int]]::new()
        [DictStatic]::Weights.Add("A", 10)
        [DictStatic]::Weights.Add("B", 20)
    }
}
if ([DictStatic]::Weights["A"] -ne 10 -or [DictStatic]::Weights["B"] -ne 20) {
    Write-Host "FAIL: Static Dictionary failed"
    exit 1
}
Write-Host "PASS"
exit 0
