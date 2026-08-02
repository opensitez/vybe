// vybe-test: csharp/exceptions_advanced/custom_exception_hierarchy
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class BaseError : Exception {
    public BaseError(string msg) : base(msg) {}
}
class NotFoundError : BaseError {
    public NotFoundError(string msg) : base(msg) {}
}
try {
    throw new NotFoundError("user missing");
} catch (BaseError e) {
    __Check(("base: " + e.Message).ToString(), "base: user missing");
}
