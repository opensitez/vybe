// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_in_switch_expression
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Mode(2).Name()).ToString());
__Check("b");

class Mode(int code) { public string Name() => code switch { 1 => "a", 2 => "b", _ => "x" }; }

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
