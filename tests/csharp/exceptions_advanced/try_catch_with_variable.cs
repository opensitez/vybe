// vybe-test: csharp/exceptions_advanced/try_catch_with_variable
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    int.Parse("notanumber");
} catch (Exception e) {
    __Check(("Error: " + e.Message).ToString(), "Error: Input string was not in a correct format.");
}
