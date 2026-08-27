// vybe-test: csharp/collections_advanced/collection_initializer_dict
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

using static __Harness;

var ages = new Dictionary<string, int> {
    { "Alice", 30 },
    { "Bob", 25 }
}
;
__P((ages["Alice"]).ToString());
__P((ages.Count).ToString());
__Check("30\n2");

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
