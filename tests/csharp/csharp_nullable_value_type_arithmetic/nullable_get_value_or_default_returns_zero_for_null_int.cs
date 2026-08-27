// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_get_value_or_default_returns_zero_for_null_int
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

int? value = null;
__P((value.GetValueOrDefault()).ToString());
__P((value.GetValueOrDefault(99)).ToString());
__Check("0\n99");

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
