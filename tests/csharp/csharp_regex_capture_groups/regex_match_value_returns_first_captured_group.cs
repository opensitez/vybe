// vybe-test: csharp/csharp_regex_capture_groups/regex_match_value_returns_first_captured_group
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

using static __Harness;

var match = System.Text.RegularExpressions.Regex.Match("id=42", @"id=(\d+)");
__P((match.Groups[1].Value).ToString());
__Check("42");

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
