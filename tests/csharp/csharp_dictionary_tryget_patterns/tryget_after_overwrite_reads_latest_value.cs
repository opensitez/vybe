// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_overwrite_reads_latest_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["k"] = 1 }
;
map["k"] = 9;
map.TryGetValue("k", out int v);
__P((v).ToString());
__Check("9");

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
