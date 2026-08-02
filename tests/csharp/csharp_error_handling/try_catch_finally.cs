// vybe-test: csharp/csharp_error_handling/try_catch_finally
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
    throw new Exception("fail");
} catch (Exception e) {
    __Check(("caught: " + e.Message).ToString(), "caught: fail");
} finally {
    __Check(("cleanup").ToString(), "cleanup");
}
