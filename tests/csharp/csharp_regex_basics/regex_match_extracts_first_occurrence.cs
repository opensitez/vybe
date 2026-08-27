// vybe-test: csharp/csharp_regex_basics/regex_match_extracts_first_occurrence
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

using static __Harness;

var m=System.Text.RegularExpressions.Regex.Match("abc123def","[0-9]+");
__P((m.Value).ToString());
__Check("123");

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
