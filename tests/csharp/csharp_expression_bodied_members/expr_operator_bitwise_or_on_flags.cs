// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_bitwise_or_on_flags
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P(((new Bits { V = 1 } | new Bits { V = 2 }).V).ToString());
__Check("3");

struct Bits { public int V; public static Bits operator |(Bits a, Bits b) => new Bits { V = a.V | b.V }; }

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
