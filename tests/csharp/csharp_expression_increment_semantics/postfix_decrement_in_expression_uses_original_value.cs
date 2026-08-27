// vybe-test: csharp/csharp_expression_increment_semantics/postfix_decrement_in_expression_uses_original_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_increment_semantics.rs

using static __Harness;

int n = 3;
int total = n-- + n;
__P((total).ToString());
__P((n).ToString());
__Check("5\n2");

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
