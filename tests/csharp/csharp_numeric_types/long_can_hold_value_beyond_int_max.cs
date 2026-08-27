// vybe-test: csharp/csharp_numeric_types/long_can_hold_value_beyond_int_max
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

long x = (long)int.MaxValue + 1;
__P((x).ToString());
__Check("2147483648");

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
