// vybe-test: csharp/exceptions_advanced/exception_message_property
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new Exception("test message");
} catch (Exception e) {
    __Check((e.Message).ToString(), "test message");
}
