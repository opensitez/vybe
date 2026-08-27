// vybe-test: csharp/exceptions_advanced/argument_null_exception
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new ArgumentNullException("param1");
}
catch (ArgumentNullException e) {
    __P((e.ParamName).ToString());
}
__Check("param1");

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
