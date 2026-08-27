// vybe-test: csharp/csharp_list_dictionary/list_addrange_empty_collection_keeps_count
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<int> { 1 }
;
list.AddRange(new int[] { });
__P((list.Count).ToString());
__Check("1");

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
