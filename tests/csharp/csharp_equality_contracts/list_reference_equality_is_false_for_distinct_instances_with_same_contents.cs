// vybe-test: csharp/csharp_equality_contracts/list_reference_equality_is_false_for_distinct_instances_with_same_contents
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

using static __Harness;
using System.Collections.Generic;
using System.Linq;

var left = new List<int> { 1, 2 }
;
var right = new List<int> { 1, 2 }
;
__P((left == right).ToString());
__P((left.SequenceEqual(right)).ToString());
__Check("False\nTrue");

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
