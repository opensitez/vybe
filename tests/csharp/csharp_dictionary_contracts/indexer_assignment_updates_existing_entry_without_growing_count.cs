// vybe-test: csharp/csharp_dictionary_contracts/indexer_assignment_updates_existing_entry_without_growing_count
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["x"] = 1 }
;
map["x"] = 9;
__P((map["x"]).ToString());
__P((map.Count).ToString());
__Check("9\n1");

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
