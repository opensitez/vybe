// vybe-test: csharp/csharp_map_set_collections/sorted_set_enumerates_numbers_in_ascending_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var set = new SortedSet<int> { 5, 1, 3 }
;
foreach (var item in set) __P((item).ToString());
__Check("1\n3\n5");

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
