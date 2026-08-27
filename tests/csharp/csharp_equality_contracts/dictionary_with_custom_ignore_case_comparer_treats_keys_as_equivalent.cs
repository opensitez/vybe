// vybe-test: csharp/csharp_equality_contracts/dictionary_with_custom_ignore_case_comparer_treats_keys_as_equivalent
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
map["User"] = 1;
__P((map.ContainsKey("user")).ToString());
__P((map["USER"]).ToString());
__Check("True\n1");

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
