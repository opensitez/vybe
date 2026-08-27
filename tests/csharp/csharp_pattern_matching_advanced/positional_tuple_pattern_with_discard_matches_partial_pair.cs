// vybe-test: csharp/csharp_pattern_matching_advanced/positional_tuple_pattern_with_discard_matches_partial_pair
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

var pair = (2, 9);
if (pair is (2, _)) __P(("left-two").ToString());
__Check("left-two");

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
