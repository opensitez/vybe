// vybe-test: csharp/csharp_switch_type_patterns/switch_expression_returns_value_from_matching_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

int n = 2;
string word = n switch { 1 => "one", 2 => "two", _ => "many" }
;
__P((word).ToString());
__Check("two");

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
