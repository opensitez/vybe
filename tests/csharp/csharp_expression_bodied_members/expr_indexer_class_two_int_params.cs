// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_two_int_params
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Grid()[1, 0]).ToString());
__Check("3");

class Grid { int[,] m = { { 1, 2 }, { 3, 4 } }; public int this[int r, int c] => m[r, c]; }

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
