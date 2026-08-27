// vybe-test: csharp/csharp_list_dictionary/list_nested_foreach_counts_each_inner_length
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using static __Harness;
using System.Collections.Generic;

var outer = new List<List<int>> { new List<int> { 1, 2 }, new List<int> { 3 } }
;
foreach (var inner in outer) __P((inner.Count).ToString());
__Check("2\n1");

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
