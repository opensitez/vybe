// vybe-test: csharp/csharp_scope_variables/out_variable_declared_inline_at_call_site
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

using static __Harness;

if(int.TryParse("42", out int n)) __P((n).ToString());
else __P((0).ToString());
__Check("42");

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
