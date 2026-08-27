// vybe-test: csharp/csharp_collection_initializer_syntax/list_initializer_populates_elements_in_source_order
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;
using System.Collections.Generic;

var items = new List<int> { 3, 1, 4 }
;
__P((items[0]).ToString());
__P((items[2]).ToString());
__Check("3\n4");

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
