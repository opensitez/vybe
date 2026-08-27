// vybe-test: csharp/csharp_switch_expression_core/switch_expr_enum_symbolic_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

var m=Mode.On;
__P((m switch{Mode.Off=>"0",Mode.On=>"1",_=>"?"}).ToString());
__Check("1");

enum Mode { Off, On }

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
