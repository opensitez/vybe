// vybe-test: csharp/csharp_hashset_set_algebra/symmetric_except_twice_restores_original_when_same_other
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

using static __Harness;
using System.Collections.Generic;

var a = new HashSet<int> { 1, 2, 3 }
;
var other = new[] { 2, 4 }
;
a.SymmetricExceptWith(other);
a.SymmetricExceptWith(other);
__P((a.SetEquals(new HashSet<int> { 1, 2, 3 })).ToString());
__Check("True");

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
