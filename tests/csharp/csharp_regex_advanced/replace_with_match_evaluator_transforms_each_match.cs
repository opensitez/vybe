// vybe-test: csharp/csharp_regex_advanced/replace_with_match_evaluator_transforms_each_match
// origin: languages/csharp/tests/csharp/test_csharp_regex_advanced.rs

using static __Harness;

string result = System.Text.RegularExpressions.Regex.Replace(
    "a1b2c3", @"\d",
    m => ((int.Parse(m.Value)*2)).ToString());
__P((result).ToString());
__Check("a2b4c6");

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
