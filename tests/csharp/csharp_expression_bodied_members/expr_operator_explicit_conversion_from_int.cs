// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_explicit_conversion_from_int
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

Wrap w = (Wrap)9;
__P((w.V).ToString());
__Check("9");

struct Wrap { public int V; public static explicit operator Wrap(int n) => new Wrap { V = n }; }

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
