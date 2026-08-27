// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_equality_compares_values_not_references
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

int? a = 7;
int? b = 7;
__P((a == b).ToString());
int? c = null;
__P((a == c).ToString());
__P((c == null).ToString());
__Check("True\nFalse\nTrue");

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
