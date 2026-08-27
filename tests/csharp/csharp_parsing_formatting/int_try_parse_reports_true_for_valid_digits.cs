// vybe-test: csharp/csharp_parsing_formatting/int_try_parse_reports_true_for_valid_digits
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

using static __Harness;

var ok = int.TryParse("42", out var value);
__P((ok).ToString());
__P((value).ToString());
__Check("True\n42");

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
