// vybe-test: csharp/csharp_string_parsing/decimal_parse_preserves_exact_fraction
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

using static __Harness;

var d=decimal.Parse("0.1",System.Globalization.CultureInfo.InvariantCulture);
__P((d+0.2m==0.3m).ToString());
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
