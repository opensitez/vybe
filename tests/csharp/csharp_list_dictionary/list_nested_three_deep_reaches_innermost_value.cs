// vybe-test: csharp/csharp_list_dictionary/list_nested_three_deep_reaches_innermost_value
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

using static __Harness;
using System.Collections.Generic;

var outer = new List<List<List<int>>>();
var mid = new List<List<int>>();
var inner = new List<int> { 5 }
;
mid.Add(inner);
outer.Add(mid);
__P((outer[0][0][0]).ToString());
__Check("5");

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
