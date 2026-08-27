// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_struct_get_set
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var p = new PairStore();
p[0] = 3;
p[1] = 9;
__P((p[0]).ToString());
__P((p[1]).ToString());
__Check("3\n9");

struct PairStore { int a, b; public int this[int slot] { get => slot == 0 ? a : b; set { if (slot == 0) a = value; else b = value; } } }

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
