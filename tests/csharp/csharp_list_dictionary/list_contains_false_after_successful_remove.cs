// vybe-test: csharp/csharp_list_dictionary/list_contains_false_after_successful_remove
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<int> { 5 }
;
list.Remove(5);
__P((list.Contains(5)).ToString());
__Check("False");

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
