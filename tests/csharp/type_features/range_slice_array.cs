// vybe-test: csharp/type_features/range_slice_array
// origin: languages/csharp/tests/csharp/test_type_features.rs

using static __Harness;

var arr = new int[] { 10, 20, 30, 40, 50 }
;
var sub = arr[1..3];
__P((sub[0]).ToString());
__P((sub[1]).ToString());
__Check("20\n30");

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
