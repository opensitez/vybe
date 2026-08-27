// vybe-test: csharp/csharp_dictionary_contracts/indexer_read_returns_stored_value_for_existing_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["pi"] = 3 }
;
__P((map["pi"]).ToString());
__Check("3");

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
