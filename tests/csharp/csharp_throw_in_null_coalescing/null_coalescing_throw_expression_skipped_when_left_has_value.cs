// vybe-test: csharp/csharp_throw_in_null_coalescing/null_coalescing_throw_expression_skipped_when_left_has_value
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

using static __Harness;

string? present = "ok";
string value = present ?? throw new System.Exception("fail");
__P((value).ToString());
__Check("ok");

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
