// vybe-test: csharp/csharp_collection_expressions/collection_expression_triple_spread_merge
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

int[] a = [1];
int[] b = [2];
int[] c = [3];
int[] all = [..a, ..b, ..c];
__P((all.Length).ToString());
__P((all[2]).ToString());
__Check("3\n3");

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
