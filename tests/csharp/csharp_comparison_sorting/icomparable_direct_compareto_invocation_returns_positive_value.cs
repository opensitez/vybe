// vybe-test: csharp/csharp_comparison_sorting/icomparable_direct_compareto_invocation_returns_positive_value
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;

var left = new Rank();
left.Value = 3;
var right = new Rank();
right.Value = 1;
__P((left.CompareTo(right)).ToString());
__Check("1");

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
