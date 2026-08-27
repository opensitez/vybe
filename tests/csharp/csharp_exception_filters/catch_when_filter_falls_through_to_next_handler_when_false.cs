// vybe-test: csharp/csharp_exception_filters/catch_when_filter_falls_through_to_next_handler_when_false
// origin: languages/csharp/tests/csharp/test_csharp_exception_filters.rs

using static __Harness;

try {
    throw new System.Exception("code=500");
}
catch (System.Exception e) when (e.Message.Contains("404")) {
    __P(("not found").ToString());
}
catch (System.Exception) {
    __P(("server error").ToString());
}
__Check("server error");

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
