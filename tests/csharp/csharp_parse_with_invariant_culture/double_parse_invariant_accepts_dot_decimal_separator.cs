// vybe-test: csharp/csharp_parse_with_invariant_culture/double_parse_invariant_accepts_dot_decimal_separator
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

using static __Harness;

double value = double.Parse("3.5", System.Globalization.CultureInfo.InvariantCulture);
__P((value).ToString());
__Check("3.5");

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
