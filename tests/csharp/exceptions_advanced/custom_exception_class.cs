// vybe-test: csharp/exceptions_advanced/custom_exception_class
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

class AppException : Exception {
    public int Code { get; set; }
    public AppException(string message, int code) : base(message) {
        Code = code;
    }
}
try {
    throw new AppException("not found", 404);
} catch (AppException e) {
    __P((e.Message + " (" + e.Code + ")").ToString());
}
__Check("not found (404)");
