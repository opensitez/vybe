// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_negative_keys_sort_numerically
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var sd = new SortedDictionary<int, string> { [-1] = "neg", [0] = "zero", [1] = "pos" }
;
int first = 0;
foreach (var k in sd.Keys) { first = k; break; }
__P((first).ToString());
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
