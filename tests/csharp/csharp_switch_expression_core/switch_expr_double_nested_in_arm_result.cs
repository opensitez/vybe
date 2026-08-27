// vybe-test: csharp/csharp_switch_expression_core/switch_expr_double_nested_in_arm_result
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

int Pick(int a,int b)=>a switch{1=>b switch{2=>10,3=>20,_=>0},_=>-1}
;
__P((Pick(1,3)).ToString());
__Check("20");

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
