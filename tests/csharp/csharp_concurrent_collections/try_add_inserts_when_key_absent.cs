// vybe-test: csharp/csharp_concurrent_collections/try_add_inserts_when_key_absent
// origin: languages/csharp/tests/csharp/test_csharp_concurrent_collections.rs

using static __Harness;

var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
__P((d.TryAdd("a", 1)).ToString());
__P((d["a"]).ToString());
__Check("True\n1");

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
