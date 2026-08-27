// vybe-test: csharp/csharp_with_expression_records/with_positional_preserves_string
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var q=(new Pair("x",1)) with{N=2}
;
__P((q.S).ToString());
__Check("x");

record Pair(string S,int N);

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
