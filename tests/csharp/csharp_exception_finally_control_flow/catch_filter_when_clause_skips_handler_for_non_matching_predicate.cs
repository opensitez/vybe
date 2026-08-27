// vybe-test: csharp/csharp_exception_finally_control_flow/catch_filter_when_clause_skips_handler_for_non_matching_predicate
// origin: languages/csharp/tests/csharp/test_csharp_exception_finally_control_flow.rs

using static __Harness;

string label = "start";
try {
    throw new Exception("code-404");
}
catch (Exception e) when (e.Message.Contains("500")) {
    label = "wrong";
}
catch (Exception e) when (e.Message.Contains("404")) {
    label = "matched";
}
__P((label).ToString());
__Check("matched");

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
