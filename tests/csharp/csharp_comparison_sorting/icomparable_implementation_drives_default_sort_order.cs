// vybe-test: csharp/csharp_comparison_sorting/icomparable_implementation_drives_default_sort_order
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;
using System.Collections.Generic;

var list = new List<Rank> { new Rank { Value = 3 }, new Rank { Value = 1 } }
;
list.Sort();
foreach (var item in list) __P((item.Value).ToString());
__Check("1\n3");

class Rank : System.IComparable<Rank> { public int Value; public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } }

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
