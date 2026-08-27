// vybe-test: csharp/csharp_index_range_slice/range_closed_start_open_end_excludes_last_element
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

int[] data={5,6,7,8,9}
;
var slice=data[1..^1];
__P((slice.Length).ToString());
__P((slice[0]).ToString());
__P((slice[1]).ToString());
__Check("3\n6\n7");

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
