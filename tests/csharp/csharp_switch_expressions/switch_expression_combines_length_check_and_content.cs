// vybe-test: csharp/csharp_switch_expressions/switch_expression_combines_length_check_and_content
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var text = "tool";
__P((text switch { string s when s.Length == 4 => "len4", string s => s, _ => "none" }).ToString());
__Check("len4");

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
