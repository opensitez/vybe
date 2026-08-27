// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_tuple_with_discard_component
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var pair = (3, 9);
__P((pair switch { (3, _) => "starts-three", (_, 9) => "ends-nine", _ => "other" }).ToString());
__Check("starts-three");

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
