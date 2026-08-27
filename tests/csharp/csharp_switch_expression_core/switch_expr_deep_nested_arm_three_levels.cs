// vybe-test: csharp/csharp_switch_expression_core/switch_expr_deep_nested_arm_three_levels
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

string L(int n)=>n switch{1=>"a",2=>2 switch{2=>"b",3=>"c",_=>"d"},_=>"z"}
;
__P((L(2)).ToString());
__Check("b");

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
