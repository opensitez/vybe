// vybe-test: csharp/csharp_index_range_slice/index_from_end_on_two_element_array
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

using static __Harness;

int[] data={7,8}
;
__P((data[^1]).ToString());
__P((data[^2]).ToString());
__Check("8\n7");

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
