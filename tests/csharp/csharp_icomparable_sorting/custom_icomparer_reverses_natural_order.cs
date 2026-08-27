// vybe-test: csharp/csharp_icomparable_sorting/custom_icomparer_reverses_natural_order
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

using static __Harness;

var list = new System.Collections.Generic.List<int>{3,1,4,1,5}
;
list.Sort(new Desc());
__P((list[0]).ToString());
__Check("5");

class Desc : System.Collections.Generic.IComparer<int> {
    public int Compare(int x, int y) => y.CompareTo(x);
}

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
