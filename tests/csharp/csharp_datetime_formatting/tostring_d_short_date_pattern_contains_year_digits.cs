// vybe-test: csharp/csharp_datetime_formatting/tostring_d_short_date_pattern_contains_year_digits
// origin: languages/csharp/tests/csharp/test_csharp_datetime_formatting.rs

using static __Harness;

var d = new System.DateTime(2025,12,31);
__P((d.ToString("yyyy-MM-dd").StartsWith("2025")).ToString());
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
