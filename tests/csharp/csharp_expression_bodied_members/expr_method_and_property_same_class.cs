// vybe-test: csharp/csharp_expression_bodied_members/expr_method_and_property_same_class
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Widget().Twice()).ToString());
__Check("10");

class Widget { public int baseVal = 5; public int Base => baseVal; public int Twice() => Base * 2; }

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
