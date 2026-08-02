// vybe-test: csharp/csharp_abstract_class/template_method_calls_abstract_hook_overridden_by_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

abstract class Report{
    protected abstract string Header();
    protected abstract string Body();
    public string Generate()=>Header()+"\n"+Body();
}
class HtmlReport:Report{
    protected override string Header()=>"<html>";
    protected override string Body()=>"<body></body>";
}
Console.WriteLine(new HtmlReport().Generate());
