// vybe-test: csharp/csharp_initialization_order/static_field_initializers_run_in_declaration_order_before_static_method
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

using static __Harness;

__P("Valid_static_field_initializers_run_in_declaration_order_before_static_method");
__Check("Valid_static_field_initializers_run_in_declaration_order_before_static_method");
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
