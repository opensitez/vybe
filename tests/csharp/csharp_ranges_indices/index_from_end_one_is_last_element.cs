// vybe-test: csharp/csharp_ranges_indices/index_from_end_one_is_last_element
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

using static __Harness;

int[] a={1,2,3,4,5}
;
__P((a[^1]).ToString());
__Check("5");

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
