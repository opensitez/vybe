// vybe-test: csharp/csharp_dictionary_tryget_patterns/indexer_read_matches_tryget_for_same_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["x"] = 4 }
;
map.TryGetValue("x", out int viaTry);
__P((map["x"] == viaTry).ToString());
__Check("True");

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
