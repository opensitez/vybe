// vybe-test: csharp/csharp_map_set_collections/dictionary_keys_collection_can_be_enumerated
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["a"] = 1, ["b"] = 2 }
;
foreach (var key in map.Keys) __P((key).ToString());
__Check("a\nb");

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
