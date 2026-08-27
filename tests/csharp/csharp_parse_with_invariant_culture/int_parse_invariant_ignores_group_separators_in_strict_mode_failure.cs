// vybe-test: csharp/csharp_parse_with_invariant_culture/int_parse_invariant_ignores_group_separators_in_strict_mode_failure
// origin: languages/csharp/tests/csharp/test_csharp_parse_with_invariant_culture.rs

using static __Harness;

try {
    int.Parse("1,234", System.Globalization.CultureInfo.InvariantCulture);
    __P(("parsed").ToString());
}
catch (System.FormatException) {
    __P(("reject").ToString());
}
__Check("reject");

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
