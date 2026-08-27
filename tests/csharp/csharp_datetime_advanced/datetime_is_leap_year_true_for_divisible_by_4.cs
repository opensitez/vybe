// vybe-test: csharp/csharp_datetime_advanced/datetime_is_leap_year_true_for_divisible_by_4
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

using static __Harness;

__P((System.DateTime.IsLeapYear(2024)).ToString());
__P((System.DateTime.IsLeapYear(2023)).ToString());
__Check("True\nFalse");

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
