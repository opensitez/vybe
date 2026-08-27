// vybe-test: csharp/csharp_exception_types/argument_null_exception_message_contains_param_name
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

using static __Harness;

try { throw new System.ArgumentNullException("value"); }
catch(System.ArgumentNullException e) { __P((e.ParamName).ToString()); }
__Check("value");

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
