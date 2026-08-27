// vybe-test: csharp/csharp_collection_expressions/collection_expression_multiple_spreads_with_literals
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

using static __Harness;

int[] a = [1, 2];
int[] b = [3];
int[] c = [0, ..a, ..b, 4];
__P((c[0]).ToString());
__P((c[4]).ToString());
__Check("0\n4");

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
