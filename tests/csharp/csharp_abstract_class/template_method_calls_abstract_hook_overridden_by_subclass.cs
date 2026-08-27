// vybe-test: csharp/csharp_abstract_class/template_method_calls_abstract_hook_overridden_by_subclass
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

using static __Harness;

__P((new HtmlReport().Generate()).ToString());
__Check("<html>\n<body></body>");

abstract class Report{
    protected abstract string Header();
    protected abstract string Body();
    public string Generate()=>Header()+"\n"+Body();
}

class HtmlReport:Report{
    protected override string Header()=>"<html>";
    protected override string Body()=>"<body></body>";
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
