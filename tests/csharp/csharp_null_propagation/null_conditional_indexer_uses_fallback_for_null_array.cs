// vybe-test: csharp/csharp_null_propagation/null_conditional_indexer_uses_fallback_for_null_array
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;

int[] values = null;
__P((values?[0] ?? -1).ToString());
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
