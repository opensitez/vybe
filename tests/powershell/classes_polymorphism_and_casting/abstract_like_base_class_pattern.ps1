# vybe-test: powershell/classes_polymorphism_and_casting/abstract_like_base_class_pattern
class AbstractWorker {
    [string]DoWork() { throw [System.NotImplementedException]::new() }
}
class ConcreteWorker : AbstractWorker {
    [string]DoWork() { return "Done" }
}
[AbstractWorker]$w = [ConcreteWorker]::new()
if ($w.DoWork() -ne "Done") {
    Write-Host "FAIL: Abstract worker pattern failed"
    exit 1
}
Write-Host "PASS"
exit 0
