// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_string_keys_sort_lexicographically
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var sd = new SortedDictionary<string, int> { ["zebra"] = 1, ["apple"] = 2, ["mango"] = 3 }
;
foreach (var k in sd.Keys) __P((k).ToString());
__Check("apple\nmango\nzebra");

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
