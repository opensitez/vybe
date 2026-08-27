// vybe-test: csharp/csharp_comparison_sorting/icomparable_compareto_after_list_indexing_returns_positive_value
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<Rank> { new Rank(3), new Rank(1) }
;
__P((list[0].CompareTo(list[1])).ToString());
__Check("1");

class Rank : System.IComparable<Rank> { public int Value; public Rank(int value) { Value = value; } public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } }

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
