// vybe-test: csharp/csharp_map_set_collections/hashset_rejects_duplicate_value_addition
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var set = new HashSet<int>();
__P((set.Add(3)).ToString());
__P((set.Add(3)).ToString());
__Check("True\nFalse");

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
