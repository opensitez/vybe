// vybe-test: csharp/csharp_type_conversions/nullable_value_type_is_pattern_extracts_underlying_number
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

using static __Harness;

int? maybe = 30;
if (maybe is int value) __P((value / 3).ToString());
__Check("10");

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
