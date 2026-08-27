// vybe-test: csharp/csharp_parse_with_invariant_culture/decimal_to_string_invariant_uses_dot_not_comma_for_fractions
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

using static __Harness;

decimal value = 2.25m;
__P((value.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString());
__Check("2.25");

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
