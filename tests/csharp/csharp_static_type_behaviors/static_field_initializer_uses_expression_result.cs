// vybe-test: csharp/csharp_static_type_behaviors/static_field_initializer_uses_expression_result
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

using static __Harness;

__P((Limits.Max).ToString());
__Check("64");

class Limits {
    public static int Max = 8 * 8;
}

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
