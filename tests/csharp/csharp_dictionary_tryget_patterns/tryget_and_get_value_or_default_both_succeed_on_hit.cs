// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_and_get_value_or_default_both_succeed_on_hit
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["z"] = 44 }
;
bool ok = map.TryGetValue("z", out int t);
int g = map.GetValueOrDefault("z");
__P((ok).ToString());
__P((g).ToString());
__Check("True\n44");

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
