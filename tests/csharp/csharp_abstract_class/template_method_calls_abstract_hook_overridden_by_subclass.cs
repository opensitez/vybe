// vybe-test: csharp/csharp_abstract_class/template_method_calls_abstract_hook_overridden_by_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

abstract class Report{
    protected abstract string Header();
    protected abstract string Body();
    public string Generate()=>Header()+"\n"+Body();
}
class HtmlReport:Report{
    protected override string Header()=>"<html>";
    protected override string Body()=>"<body></body>";
}
__P((new HtmlReport().Generate()).ToString());
__Check("<html>\n<body></body>");
