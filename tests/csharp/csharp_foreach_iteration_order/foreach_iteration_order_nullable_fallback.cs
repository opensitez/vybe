// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

using static __Harness;

// foreach_iteration_order
int? maybe = null;
int fallback = maybe ?? 46;
__P((fallback == 46).ToString());
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
