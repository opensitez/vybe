// vybe-test: csharp/csharp_collection_initializer_syntax/list_initializer_after_empty_constructor_appends_in_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;
using System.Collections.Generic;

var items = new List<string>();
items.Add("first");
items.Add("second");
__P((items[1]).ToString());
__Check("second");

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
