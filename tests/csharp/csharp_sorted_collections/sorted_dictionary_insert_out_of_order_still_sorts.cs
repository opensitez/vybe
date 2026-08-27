// vybe-test: csharp/csharp_sorted_collections/sorted_dictionary_insert_out_of_order_still_sorts
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var sd = new SortedDictionary<int, int>();
sd[30] = 3;
sd[10] = 1;
sd[20] = 2;
int sum = 0;
foreach (var p in sd) sum += p.Key;
__P((sum).ToString());
__Check("60");

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
