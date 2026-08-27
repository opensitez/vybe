// vybe-test: csharp/csharp_ranges_indices/range_open_end_takes_suffix
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

using static __Harness;

int[] a={1,2,3,4,5}
;
var s=a[3..];
__P((s.Length).ToString());
__P((s[0]).ToString());
__Check("2\n4");

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
