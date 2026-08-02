// vybe-test: csharp/exceptions_advanced/try_finally_always_runs
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    __Check(("in try").ToString(), "in try");
} finally {
    __Check(("in finally").ToString(), "in finally");
}
