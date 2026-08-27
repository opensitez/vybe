// vybe-test: csharp/csharp_string_parsing/datetime_try_parse_exact_with_format
// origin: languages/csharp/tests/csharp/test_csharp_string_parsing.rs

using static __Harness;

bool ok=System.DateTime.TryParseExact("2024-01-15","yyyy-MM-dd",
    System.Globalization.CultureInfo.InvariantCulture,
    System.Globalization.DateTimeStyles.None,out var dt);
__P((ok).ToString());
__P((dt.Day).ToString());
__Check("True\n15");

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
