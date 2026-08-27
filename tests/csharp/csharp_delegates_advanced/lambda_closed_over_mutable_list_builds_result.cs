// vybe-test: csharp/csharp_delegates_advanced/lambda_closed_over_mutable_list_builds_result
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

using static __Harness;

Action act = () => __P("Delegated");
act();
__Check("Delegated");
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
