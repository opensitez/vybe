// vybe-test: csharp/csharp_regex_capture_groups/regex_options_ignore_case_matches_differing_casing
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

using static __Harness;

bool ok = System.Text.RegularExpressions.Regex.IsMatch(
    "Hello",
    "hello",
    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
__P((ok).ToString());
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
