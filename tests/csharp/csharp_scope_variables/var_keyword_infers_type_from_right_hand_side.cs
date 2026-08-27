// vybe-test: csharp/csharp_scope_variables/var_keyword_infers_type_from_right_hand_side
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

using static __Harness;

var text = "hello";
var number = 42;
__P((text.GetType().Name).ToString());
__P((number.GetType().Name).ToString());
__Check("String\nInt32");

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
