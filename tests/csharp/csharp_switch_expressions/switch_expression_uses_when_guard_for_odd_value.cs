// vybe-test: csharp/csharp_switch_expressions/switch_expression_uses_when_guard_for_odd_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var x = 5;
__P((x switch { int n when n % 2 == 0 => "even", int n => "odd" }).ToString());
__Check("odd");

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
