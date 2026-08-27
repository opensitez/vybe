// vybe-test: csharp/csharp_hashset_set_algebra/symmetric_except_with_keeps_elements_in_exactly_one_set
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

using static __Harness;
using System.Collections.Generic;

var a = new HashSet<int> { 1, 2, 3 }
;
a.SymmetricExceptWith(new[] { 2, 3, 4 });
__P((a.Contains(1)).ToString());
__P((a.Contains(4)).ToString());
__P((a.Contains(2)).ToString());
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
