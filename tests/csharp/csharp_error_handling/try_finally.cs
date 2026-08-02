// vybe-test: csharp/csharp_error_handling/try_finally
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    __Check(("try").ToString(), "try");
} finally {
    __Check(("finally").ToString(), "finally");
}
