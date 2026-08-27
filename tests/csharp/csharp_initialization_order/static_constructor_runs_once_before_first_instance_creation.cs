// vybe-test: csharp/csharp_initialization_order/static_constructor_runs_once_before_first_instance_creation
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

using static __Harness;

new Counter();
new Counter();
__Check("static-ctor\ninstance\ninstance");

class Counter {
    static Counter() { __P(("static-ctor").ToString()); }
    public Counter() { __P(("instance").ToString()); }
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
