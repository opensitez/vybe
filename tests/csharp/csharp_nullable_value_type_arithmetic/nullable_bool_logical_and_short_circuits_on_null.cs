// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_bool_logical_and_short_circuits_on_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

using static __Harness;

bool? t = true;
bool? n = null;
bool? f = false;
__P((t & n).ToString());
__P((n & f).ToString());
__P((f & t).ToString());
__Check("\nFalse\nFalse");

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
