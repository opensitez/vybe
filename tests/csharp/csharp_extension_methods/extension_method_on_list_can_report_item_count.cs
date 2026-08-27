// vybe-test: csharp/csharp_extension_methods/extension_method_on_list_can_report_item_count
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

using static __Harness;
using Demo;
using System.Collections.Generic;

__P((new List<int> { 1, 2 }.Describe()).ToString());
__Check("count=2");

namespace Demo { public static class ListExt { public static string Describe<T>(this List<T> values) { return "count=" + values.Count; } } }

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
