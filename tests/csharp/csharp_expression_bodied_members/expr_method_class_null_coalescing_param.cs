// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_null_coalescing_param
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Safe().OrEmpty(null)).ToString());
__P((new Safe().OrEmpty("x")).ToString());
__Check("\nx");

class Safe { public string OrEmpty(string? s) => s ?? ""; }

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
