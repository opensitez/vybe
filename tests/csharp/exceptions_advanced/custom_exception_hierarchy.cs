// vybe-test: csharp/exceptions_advanced/custom_exception_hierarchy
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
    __P(("base: " + e.Message).ToString());
}
__Check("base: user missing");
