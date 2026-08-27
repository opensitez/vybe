// vybe-test: csharp/csharp_collection_expressions/collection_expression_spread_merges_two_arrays
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

int[] a = [1, 2, 3];
int[] b = [4, 5, 6];
int[] c = [..a, ..b];
__P((c.Length).ToString());
__P((c[3]).ToString());
__Check("6\n4");

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
