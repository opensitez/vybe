// vybe-test: csharp/exceptions_advanced/try_catch_basic
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    int x = 10 / 0;
} catch (DivideByZeroException) {
    __Check(("caught divide by zero").ToString(), "caught divide by zero");
}
