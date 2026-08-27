// vybe-test: csharp/exceptions_advanced/custom_exception_class
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new AppException("not found", 404);
}
catch (AppException e) {
    __P((e.Message + " (" + e.Code + ")").ToString());
}
__Check("not found (404)");

class AppException : Exception {
    public int Code { get; set; }
    public AppException(string message, int code) : base(message) {
        Code = code;
    }
}

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
