// vybe-test: csharp/csharp_switch_expression_core/switch_expr_selector_is_method_call
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

int Twice(int x)=>x*2;
__P((Twice(3) switch{6=>"ok",_=>"no"}).ToString());
__Check("ok");

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
