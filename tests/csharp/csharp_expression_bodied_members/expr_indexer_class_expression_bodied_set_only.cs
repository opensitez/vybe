// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_expression_bodied_set_only
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var s = new Store();
s.Value = 11;
__P((s.Value).ToString());
__Check("11");

class Store { int v; public int Value { get { return v; } set => v = value; } }

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
