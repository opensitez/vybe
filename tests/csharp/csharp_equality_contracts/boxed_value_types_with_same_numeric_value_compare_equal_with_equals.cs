// vybe-test: csharp/csharp_equality_contracts/boxed_value_types_with_same_numeric_value_compare_equal_with_equals
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

using static __Harness;

object left = 42;
object right = 42;
__P((left.Equals(right)).ToString());
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
