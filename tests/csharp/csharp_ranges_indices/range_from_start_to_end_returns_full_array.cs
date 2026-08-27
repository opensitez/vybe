// vybe-test: csharp/csharp_ranges_indices/range_from_start_to_end_returns_full_array
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

using static __Harness;

int[] a={1,2,3}
;
var s=a[..];
__P((s.Length).ToString());
__Check("3");

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
