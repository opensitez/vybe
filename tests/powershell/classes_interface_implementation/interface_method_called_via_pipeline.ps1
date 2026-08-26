# vybe-test: powershell/classes_interface_implementation/interface_method_called_via_pipeline
class PipelineWorker : System.IDisposable {
    [string]$Id
    PipelineWorker([string]$i) { $this.Id = $i }
    [void]Dispose() {}
}
$workers = @([PipelineWorker]::new("W1"), [PipelineWorker]::new("W2"))
$workers | ForEach-Object { [System.IDisposable]$_ } | ForEach-Object { $_.Dispose() }
Write-Host "PASS"
exit 0
