// vybe-test: csharp/csharp_indexer_get_set/indexer_get_set_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_indexer_get_set.rs

using static __Harness;

// indexer_get_set
string feature = "indexer_get_set";
__P((feature.Contains("a") || !feature.Contains("a")).ToString());
__Check("True");

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
