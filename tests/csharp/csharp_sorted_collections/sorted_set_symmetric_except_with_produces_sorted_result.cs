// vybe-test: csharp/csharp_sorted_collections/sorted_set_symmetric_except_with_produces_sorted_result
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var ss = new SortedSet<int> { 1, 2, 3 }
;
ss.SymmetricExceptWith(new[] { 2, 3, 4 });
__P((ss.Contains(1)).ToString());
__P((ss.Contains(4)).ToString());
__P((ss.Contains(2)).ToString());
__Check("True\nTrue\nFalse");

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
