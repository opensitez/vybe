// vybe-test: csharp/csharp_expression_bodied/expression_bodied_method_returns_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

__P((new Calc().Double(5)).ToString());
__Check("10");

class Calc{public int Double(int n)=>n*2;}

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
