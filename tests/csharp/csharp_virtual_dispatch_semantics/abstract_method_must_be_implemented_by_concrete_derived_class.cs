// vybe-test: csharp/csharp_virtual_dispatch_semantics/abstract_method_must_be_implemented_by_concrete_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

Parser parser = new EchoParser();
__P((parser.Parse("  hi  ")).ToString());
__Check("hi");

abstract class Parser {
    public abstract string Parse(string input);
}

class EchoParser : Parser {
    public override string Parse(string input) { return input.Trim(); }
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
