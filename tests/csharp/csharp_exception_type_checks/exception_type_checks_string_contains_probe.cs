// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

using static __Harness;

// exception_type_checks
string feature = "exception_type_checks";
__P((feature.Contains("a") || !feature.Contains("a")).ToString());
__Check("True");

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
