// vybe-test: csharp/csharp_null_handling/null_conditional_indexer_returns_null_on_null_collection
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

using static __Harness;

int[] arr = null;
__P((arr?[0] ?? -1).ToString());
__Check("-1");

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
