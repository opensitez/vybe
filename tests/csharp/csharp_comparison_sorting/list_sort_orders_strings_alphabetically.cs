// vybe-test: csharp/csharp_comparison_sorting/list_sort_orders_strings_alphabetically
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<string> { "c", "a", "b" }
;
list.Sort();
foreach (var value in list) __P((value).ToString());
__Check("a\nb\nc");

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
