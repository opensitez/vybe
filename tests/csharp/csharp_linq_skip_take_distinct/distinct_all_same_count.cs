// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_all_same_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

using static __Harness;

var r=new[]{7,7,7,7}
.Distinct();
__P((r.Count()).ToString());
__Check("1");

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
