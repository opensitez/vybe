// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_indexer_insert_finds_new_entry
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int>();
map["fresh"] = 42;
__P((map.TryGetValue("fresh", out int v)).ToString());
__P((v).ToString());
__Check("True\n42");

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
