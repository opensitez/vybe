// vybe-test: csharp/csharp_numeric_types/uint_holds_value_beyond_int_max
// origin: languages/csharp/tests/csharp/test_csharp_numeric_types.rs

using static __Harness;

uint u = 3000000000u;
__P((u > int.MaxValue).ToString());
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
