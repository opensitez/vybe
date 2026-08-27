// vybe-test: csharp/csharp_string_culture/invariant_culture_tostring_for_double_uses_dot_separator
// origin: languages/csharp/tests/csharp/test_csharp_string_culture.rs

using static __Harness;

double d=1.5;
__P((d.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToString());
__Check("1.5");

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
