// vybe-test: csharp/csharp_list_dictionary/dictionary_string_key_can_be_read_twice
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using static __Harness;
using System.Collections.Generic;

var map = new Dictionary<string, int> { ["k"] = 3 }
;
__P((map["k"]).ToString());
__P((map["k"]).ToString());
__Check("3\n3");

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
