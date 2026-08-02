// vybe-test: csharp/exceptions_advanced/catch_finally_on_error
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "start";
try {
    int x = 10 / 0;
    result = "never";
} catch (DivideByZeroException) {
    result = "caught";
} finally {
    result += " + finally";
}
__Check((result).ToString(), "caught + finally");
