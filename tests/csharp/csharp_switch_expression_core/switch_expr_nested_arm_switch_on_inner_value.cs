// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_arm_switch_on_inner_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

string Outer(int n)=>n switch{1=>1 switch{1=>"one-one",_=>"one-other"},2=>"two",_=>"rest"}
;
__P((Outer(1)).ToString());
__Check("one-one");

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
