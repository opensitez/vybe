// vybe-test: csharp/csharp_map_set_collections/sorted_dictionary_enumerates_keys_in_sorted_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var map = new SortedDictionary<string, int> { ["b"] = 2, ["a"] = 1 }
;
foreach (var pair in map) __P((pair.Key + ":" + pair.Value).ToString());
__Check("a:1\nb:2");

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
