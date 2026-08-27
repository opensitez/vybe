// vybe-test: csharp/csharp_throw_in_null_coalescing/null_coalescing_throw_expression_runs_when_left_is_null
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

using static __Harness;

string? missing = null;
try {
    string value = missing ?? throw new System.Exception("required");
    __P((value).ToString());
}
catch (System.Exception e) {
    __P((e.Message).ToString());
}
__Check("required");

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
