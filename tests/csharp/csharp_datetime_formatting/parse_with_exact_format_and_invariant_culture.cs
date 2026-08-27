// vybe-test: csharp/csharp_datetime_formatting/parse_with_exact_format_and_invariant_culture
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

using static __Harness;

var d = System.DateTime.ParseExact("2024-03-21","yyyy-MM-dd",
    System.Globalization.CultureInfo.InvariantCulture);
__P((d.Year).ToString());
__P((d.Month).ToString());
__P((d.Day).ToString());
__Check("2024\n3\n21");

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
