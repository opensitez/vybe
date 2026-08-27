// vybe-test: csharp/csharp_index_range_slice/array_slice_assign_is_independent_copy
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

int[] data={1,2,3}
;
var slice=data[0..2];
slice[0]=9;
__P((data[0]).ToString());
__P((slice[0]).ToString());
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
