// vybe-test: csharp/exceptions_advanced/rethrow_with_throw
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    try {
        throw new Exception("inner");
    } catch (Exception) {
        throw;
    }
} catch (Exception e) {
    __Check(("outer: " + e.Message).ToString(), "outer: inner");
}
