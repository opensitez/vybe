// vybe-test: csharp/csharp_regex_advanced/matches_returns_all_non_overlapping_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

using static __Harness;

var matches = System.Text.RegularExpressions.Regex.Matches("a1 b2 c3", @"\d");
__P((matches.Count).ToString());
__Check("3");

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
