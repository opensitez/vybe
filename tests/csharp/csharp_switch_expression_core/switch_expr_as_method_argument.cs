// vybe-test: csharp/csharp_switch_expression_core/switch_expr_as_method_argument
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

void Show(string s){__P((s).ToString());}
Show(3 switch{3=>"three",_=>"other"});
__Check("three");

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
