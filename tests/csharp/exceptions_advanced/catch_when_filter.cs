// vybe-test: csharp/exceptions_advanced/catch_when_filter
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new Exception("error 42");
} catch (Exception e) when (e.Message.Contains("42")) {
    __Check(("filtered catch: " + e.Message).ToString(), "filtered catch: error 42");
}
