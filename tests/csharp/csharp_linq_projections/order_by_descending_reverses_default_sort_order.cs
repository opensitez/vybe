// vybe-test: csharp/csharp_linq_projections/order_by_descending_reverses_default_sort_order
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var result = new[]{3,1,4,1,5}
.OrderByDescending(x => x).Distinct();
foreach(var n in result) __P((n).ToString());
__Check("5\n4\n3\n1");

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
