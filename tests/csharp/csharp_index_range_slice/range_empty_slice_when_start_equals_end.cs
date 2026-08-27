// vybe-test: csharp/csharp_index_range_slice/range_empty_slice_when_start_equals_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

int[] data={1,2,3}
;
var slice=data[2..2];
__P((slice.Length).ToString());
__Check("0");

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
