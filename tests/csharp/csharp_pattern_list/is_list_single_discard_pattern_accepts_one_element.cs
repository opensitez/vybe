// vybe-test: csharp/csharp_pattern_list/is_list_single_discard_pattern_accepts_one_element
// origin: languages/csharp/tests/csharp/test_csharp_pattern_list.rs

using static __Harness;

int[] data = new[]{7}
;
__P((data is [_]).ToString());
__Check("True");

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
