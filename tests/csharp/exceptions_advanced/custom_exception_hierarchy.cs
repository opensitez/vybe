// vybe-test: csharp/exceptions_advanced/custom_exception_hierarchy
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

using static __Harness;

try {
    throw new NotFoundError("user missing");
}
catch (BaseError e) {
    __P(("base: " + e.Message).ToString());
}
__Check("base: user missing");

class BaseError : Exception {
    public BaseError(string msg) : base(msg) {}
}

class NotFoundError : BaseError {
    public NotFoundError(string msg) : base(msg) {}
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
