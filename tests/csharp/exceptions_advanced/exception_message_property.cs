// vybe-test: csharp/exceptions_advanced/exception_message_property
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new Exception("test message");
}
catch (Exception e) {
    __P((e.Message).ToString());
}
__Check("test message");

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
