// vybe-test: csharp/csharp_pattern_deconstruct/list_pattern_with_slice_matches_prefix_and_suffix
// origin: languages/csharp/tests/csharp/test_csharp_pattern_deconstruct.rs

using static __Harness;

int[] data = { 1, 2, 3, 4, 5 }
;
if (data is [1, .., 5]) __P(("bookended").ToString());
else __P(("no").ToString());
__Check("bookended");

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
