// vybe-test: csharp/exceptions_advanced/catch_hierarchy_most_specific_first
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new InvalidOperationException("invalid op");
}
catch (InvalidOperationException e) {
    __P(("specific: " + e.Message).ToString());
}
catch (Exception) {
    __P(("generic").ToString());
}
__Check("specific: invalid op");

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
