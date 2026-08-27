// vybe-test: csharp/csharp_numeric_types/int_min_value_is_negative_two_billion_one_forty_eight_million
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

__P((int.MinValue).ToString());
__Check("-2147483648");

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
