// vybe-test: csharp/csharp_dictionary_contracts/try_get_value_reports_absence_without_adding_entries
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["a"] = 1 }
;
bool found = map.TryGetValue("missing", out var value);
__P((found ? "Y" : "N").ToString());
__P((map.Count).ToString());
__Check("N\n1");

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
