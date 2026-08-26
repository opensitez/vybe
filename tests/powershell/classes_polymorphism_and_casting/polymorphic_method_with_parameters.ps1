# vybe-test: powershell/classes_polymorphism_and_casting/polymorphic_method_with_parameters
class FormatterBase {
    [string]Wrap([string]$msg) { return "[$msg]" }
}
class HtmlFormatter : FormatterBase {
    [string]Wrap([string]$msg) { return "<b>$msg</b>" }
}
[FormatterBase]$fmt = [HtmlFormatter]::new()
if ($fmt.Wrap("text") -ne "<b>text</b>") {
    Write-Host "FAIL: Polymorphic method with params failed"
    exit 1
}
Write-Host "PASS"
exit 0
