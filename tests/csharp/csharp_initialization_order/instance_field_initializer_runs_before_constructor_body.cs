// vybe-test: csharp/csharp_initialization_order/instance_field_initializer_runs_before_constructor_body
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

using static __Harness;

new Widget();
__Check("field\nctor");

class Widget {
    string label = Init("field");
    public Widget() {
        __P(("ctor").ToString());
    }
    static string Init(string part) {
        __P((part).ToString());
        return part;
    }
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
