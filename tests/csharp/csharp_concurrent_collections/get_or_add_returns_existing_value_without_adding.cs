// vybe-test: csharp/csharp_concurrent_collections/get_or_add_returns_existing_value_without_adding
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

using static __Harness;

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["x"] = 5;
__P((d.GetOrAdd("x", 99)).ToString());
__Check("5");

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
