// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_and_relational_combo
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

var s = new StructIndexer(10);
__P(s[0].ToString());
__Check("10");

struct StructIndexer {
    private int val;
    public StructIndexer(int v) => val = v;
    public int this[int i] => val;
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
