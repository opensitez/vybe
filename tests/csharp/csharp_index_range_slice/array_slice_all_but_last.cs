// vybe-test: csharp/csharp_index_range_slice/array_slice_all_but_last
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

int[] data={9,8,7,6}
;
var slice=data[..^1];
__P((string.Join(",",slice)).ToString());
__Check("9,8,7");

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
