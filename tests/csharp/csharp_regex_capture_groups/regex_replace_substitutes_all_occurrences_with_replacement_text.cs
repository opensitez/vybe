// vybe-test: csharp/csharp_regex_capture_groups/regex_replace_substitutes_all_occurrences_with_replacement_text
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

using static __Harness;

var text = System.Text.RegularExpressions.Regex.Replace("a-b-c", "-", "_");
__P((text).ToString());
__Check("a_b_c");

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
