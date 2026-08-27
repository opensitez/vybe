// vybe-test: csharp/csharp_new_features/range_from_end_produces_slice_of_array
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

using static __Harness;

int[] arr = {1,2,3,4,5}
;
var last2 = arr[^2..];
__P((last2[0]).ToString());
__P((last2[1]).ToString());
__Check("4\n5");

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
