// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_conversion_from_literal_wraps_value_type
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

int? boxed = 42;
__P((boxed is int).ToString());
__P(((int)boxed).ToString());
__Check("True\n42");

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
