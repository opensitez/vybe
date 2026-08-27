// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_coalesce_prefers_left_when_has_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

int? left = 8;
__P((left ?? 100).ToString());
__Check("8");

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
