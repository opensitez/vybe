// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_out_string_default_is_null_on_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<int, string>();
map.TryGetValue(1, out string s);
__P((s == null).ToString());
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
