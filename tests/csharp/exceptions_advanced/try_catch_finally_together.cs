// vybe-test: csharp/exceptions_advanced/try_catch_finally_together
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new Exception("boom");
} catch (Exception e) {
    __Check(("caught: " + e.Message).ToString(), "caught: boom");
} finally {
    __Check(("finally").ToString(), "finally");
}
