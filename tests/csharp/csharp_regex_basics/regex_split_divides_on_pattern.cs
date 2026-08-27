// vybe-test: csharp/csharp_regex_basics/regex_split_divides_on_pattern
// origin: languages/csharp/tests/csharp/test_csharp_regex_basics.rs

using static __Harness;

var parts=System.Text.RegularExpressions.Regex.Split("one1two2three","[0-9]");
__P((parts.Length).ToString());
__P((parts[1]).ToString());
__Check("3\ntwo");

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
