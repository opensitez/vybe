// vybe-test: csharp/csharp_event_advanced/null_conditional_event_invoke_safe_when_no_subscribers
// origin: languages/csharp/tests/csharp/test_csharp_event_advanced.rs

using static __Harness;

__P("Valid_null_conditional_event_invoke_safe_when_no_subscribers");
__Check("Valid_null_conditional_event_invoke_safe_when_no_subscribers");
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
