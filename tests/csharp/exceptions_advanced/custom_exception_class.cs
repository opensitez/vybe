// vybe-test: csharp/exceptions_advanced/custom_exception_class
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
    __Check((e.Message + " (" + e.Code + ")").ToString(), "not found (404)");
}
