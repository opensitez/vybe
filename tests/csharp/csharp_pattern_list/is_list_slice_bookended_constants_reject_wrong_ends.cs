// vybe-test: csharp/csharp_pattern_list/is_list_slice_bookended_constants_reject_wrong_ends
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

int[] data = new[]{1,2,3,4,6}
;
__P((data is [1,..,5]).ToString());
__Check("False");

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
