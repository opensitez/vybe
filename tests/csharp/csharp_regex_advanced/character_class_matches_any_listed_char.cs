// vybe-test: csharp/csharp_regex_advanced/character_class_matches_any_listed_char
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

using static __Harness;

var m = System.Text.RegularExpressions.Regex.Match("hello", @"[aeiou]");
__P((m.Value).ToString());
__Check("e");

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
