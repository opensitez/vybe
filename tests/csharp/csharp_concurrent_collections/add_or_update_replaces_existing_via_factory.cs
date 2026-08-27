// vybe-test: csharp/csharp_concurrent_collections/add_or_update_replaces_existing_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

using static __Harness;

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["k"] = 1;
d.AddOrUpdate("k", 0, (key, old) => old + 10);
__P((d["k"]).ToString());
__Check("11");

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
