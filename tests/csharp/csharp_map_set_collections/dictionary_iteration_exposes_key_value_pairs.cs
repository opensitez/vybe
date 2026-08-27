// vybe-test: csharp/csharp_map_set_collections/dictionary_iteration_exposes_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["b"] = 2, ["a"] = 1 }
;
foreach (var pair in map) __P((pair.Key + ":" + pair.Value).ToString());
__Check("b:2\na:1");

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
