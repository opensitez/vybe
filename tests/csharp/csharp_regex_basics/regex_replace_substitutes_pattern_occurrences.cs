// vybe-test: csharp/csharp_regex_basics/regex_replace_substitutes_pattern_occurrences
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

using static __Harness;

string r=System.Text.RegularExpressions.Regex.Replace("a1b2c3","[0-9]","#");
__P((r).ToString());
__Check("a#b#c#");

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
