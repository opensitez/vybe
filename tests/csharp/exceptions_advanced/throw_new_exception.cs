// vybe-test: csharp/exceptions_advanced/throw_new_exception
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new InvalidOperationException("not allowed");
} catch (InvalidOperationException e) {
    __Check((e.Message).ToString(), "not allowed");
}
