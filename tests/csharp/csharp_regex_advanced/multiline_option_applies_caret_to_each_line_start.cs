// vybe-test: csharp/csharp_regex_advanced/multiline_option_applies_caret_to_each_line_start
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

using static __Harness;

var matches = System.Text.RegularExpressions.Regex.Matches(
    "start\nnew line", @"^[a-z]",
    System.Text.RegularExpressions.RegexOptions.Multiline);
__P((matches.Count).ToString());
__Check("2");

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
