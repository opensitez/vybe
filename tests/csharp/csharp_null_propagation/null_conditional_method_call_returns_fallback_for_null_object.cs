// vybe-test: csharp/csharp_null_propagation/null_conditional_method_call_returns_fallback_for_null_object
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;

string text = null;
__P((text?.ToUpper() ?? "empty").ToString());
__Check("empty");

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
