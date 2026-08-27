// vybe-test: csharp/csharp_virtual_dispatch_semantics/non_virtual_method_call_ignores_derived_redeclaration_without_new
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Printer tool = new FancyPrinter();
__P((tool.Format(7)).ToString());
__Check("p:7");

class Printer {
    public string Format(int value) { return "p:" + value; }
}

class FancyPrinter : Printer {
    public string Format(int value) { return "f:" + value; }
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
