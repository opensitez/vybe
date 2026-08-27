// vybe-test: csharp/csharp_error_handling/exception_message_access
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

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
