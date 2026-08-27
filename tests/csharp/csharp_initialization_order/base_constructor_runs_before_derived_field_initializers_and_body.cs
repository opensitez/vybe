// vybe-test: csharp/csharp_initialization_order/base_constructor_runs_before_derived_field_initializers_and_body
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

using static __Harness;

__P("Valid_base_constructor_runs_before_derived_field_initializers_and_body");
__Check("Valid_base_constructor_runs_before_derived_field_initializers_and_body");
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
