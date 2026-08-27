// vybe-test: csharp/csharp_tuples_advanced/tuple_returned_from_method_and_destructured_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

using static __Harness;

(int Min, int Max) Bounds(int[] arr) =>
    (arr.Min(), arr.Max());
var (lo, hi) = Bounds(new[]{5,1,9,3});
__P((lo).ToString());
__P((hi).ToString());
__Check("1\n9");

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
