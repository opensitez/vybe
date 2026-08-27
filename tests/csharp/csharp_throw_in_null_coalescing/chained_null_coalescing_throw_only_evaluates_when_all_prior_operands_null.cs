// vybe-test: csharp/csharp_throw_in_null_coalescing/chained_null_coalescing_throw_only_evaluates_when_all_prior_operands_null
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

using static __Harness;

string? a = null;
string? b = null;
try {
    string value = a ?? b ?? throw new System.Exception("both-null");
    __P((value).ToString());
}
catch (System.Exception) {
    __P(("caught").ToString());
}
__Check("caught");

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
