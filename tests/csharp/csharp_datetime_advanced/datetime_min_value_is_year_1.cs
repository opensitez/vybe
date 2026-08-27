// vybe-test: csharp/csharp_datetime_advanced/datetime_min_value_is_year_1
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

using static __Harness;

__P((System.DateTime.MinValue.Year).ToString());
__Check("1");

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
