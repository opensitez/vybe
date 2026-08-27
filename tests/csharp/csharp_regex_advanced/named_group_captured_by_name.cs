// vybe-test: csharp/csharp_regex_advanced/named_group_captured_by_name
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

using static __Harness;

var m = System.Text.RegularExpressions.Regex.Match("date=2024-06-15", @"(?<year>\d{4})-(?<month>\d{2})");
__P((m.Groups["year"].Value).ToString());
__P((m.Groups["month"].Value).ToString());
__Check("2024\n06");

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
