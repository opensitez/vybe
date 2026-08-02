// vybe-test: csharp/csharp_error_handling/exception_message_access
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

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
