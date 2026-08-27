// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_increment_operator_on_null_stays_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

int? value = null;
value++;
__P((value.HasValue).ToString());
__Check("False");

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
